Attribute VB_Name = "Facturar"
Option Explicit


Sub AlbaransProgramacioHITRS(data As Date)
    Dim rsInc As rdoResultset
    Dim rsIncHist As rdoResultset
    Dim fFinReparacion As Date
    Dim txtInc As String
    Dim tServit As String
    Dim mes As Integer, anyo As Integer
    
    If Month(data) = 1 Then
        mes = 12
        anyo = Year(data) - 1
    Else
        mes = Month(data) - 1
        anyo = Year(data)
    End If
    
    'INCIDENCIAS DE ANA CERRADAS EL MES ANTERIOR
    Set rsInc = Db.OpenResultset("select * from incidencias where tecnico=50 and prioridad=3 and year(ffinreparacion)=" & anyo & " and month(ffinreparacion)=" & mes & " order by ffinreparacion")
    While Not rsInc.EOF
        fFinReparacion = rsInc("FFinReparacion")
        txtInc = rsInc("id") & ": "
        
        Set rsIncHist = Db.OpenResultset("select * from Inc_Historico where id=" & rsInc("id") & " and tipo='TEXTO' order by timestamp")
        While Not rsIncHist.EOF
            txtInc = txtInc & Trim(Replace(Replace(rsIncHist("incidencia"), "<p>", ""), "</p>", ""))
            rsIncHist.MoveNext
        Wend
        txtInc = Left(txtInc, 255)
        
        tServit = "SERVIT-" & Right(Year(fFinReparacion), 2) & "-" & Right("00" & Month(fFinReparacion), 2) & "-" & Right("00" & Day(fFinReparacion), 2)
        If Not ExisteixTaula(tServit) Then CreaTaulaServit2 tServit

        '1861 es el producto de Horas de programación
        ExecutaComandaSql "INSERT into [" & tServit & "] values (newid(), getdate(), '', " & rsInc("cliente") & ", 1861, 1861, 'Inicial', 'Equip 1', 1, 0, 1, '', 91, 2, '" & txtInc & "', '', '', '', '', '')"
        
        rsInc.MoveNext
    Wend

End Sub

Sub facturaAutomatica(data As Date, emisor As String, articles As String)
    Dim client As Double, comentario As String, empCliente As String
    Dim rsEmpresas As rdoResultset, rsTiendas As rdoResultset, rsClient As rdoResultset, rsNomEmp As rdoResultset
    
    Set rsEmpresas = Db.OpenResultset("select * from constantsempresa where camp like '%CampNif%' and isnull(valor, '')<>'' order by camp")
    While Not rsEmpresas.EOF
    
        Set rsTiendas = Db.OpenResultset("select * from clients where nif='" & rsEmpresas("valor") & "' and codi in (select valor1 from paramshw)")
        If Not rsTiendas.EOF Then
            'ALBARÁN ---------------------------------------------------------------------------------------------------------------------
            'Buscar el cliente
            empCliente = ""
            If InStr(rsEmpresas("camp"), "_") Then empCliente = Split(rsEmpresas("camp"), "_")(0) & "_"
            
            Set rsNomEmp = Db.OpenResultset("select * from constantsempresa where camp = '" & empCliente & "CampNom' and isnull(valor, '')<>'' ")
            If Not rsNomEmp.EOF Then
                Set rsClient = Db.OpenResultset("select * from clients where nif='" & rsEmpresas("valor") & "' and nom = '" & rsNomEmp("valor") & "' and codi not in (select valor1 from paramshw)")
                If Not rsClient.EOF Then
                    client = CDbl(rsClient("codi"))
                    While Not rsTiendas.EOF
                        comentario = "[Lote:" & rsTiendas("nom") & "][Repercutido:" & rsTiendas("codi") & "]"
                            
                        ExecutaComandaSql "INSERT INTO [" & DonamNomTaulaServit(data) & "] (id, timestamp, client, codiarticle, viatge, Equip, quantitatDemanada, QuantitatTornada, QuantitatServida, Hora, Tipuscomanda, Comentari) values (newid(), getdate(), " & client & ", " & articles & ", 'SEMANAL', 'SOLUCIONES', 1, 0, 1, 91, 2, '" & comentario & "')"
                        
                        rsTiendas.MoveNext
                    Wend
                    
                    'FACTURA ---------------------------------------------------------------------------------------------------------------------
                    FacturaClientSemanal client, data, data, data, Format(data, "dd-mm-yyyy"), "Live Preus Actuals", " Codiarticle in (" & articles & ") and ", emisor
                End If
            End If
        End If
        rsEmpresas.MoveNext
    Wend
End Sub

Sub facturaAutomaticaSoluciones(data As Date)
    Dim client As Double, emisor As String, articulo As String
    Dim rsTiendas As rdoResultset
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'HORAS PROGRAMACIÓN ------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    client = 1253 'SILEMA
    emisor = "19" 'SOLUCIONES
    articulo = "8182" 'Horas Desarrollo Soluciones
    
    'ALBARÁN ---------------------------------------------------------------------------------------------------------------------
    ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 7.5, '', 91, 2, 'SILEMA BCN S.L. (CONTABILIDAD)', '', null, '', '', '')"
    ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 7.5, '', 91, 2, 'SILEMA BCN S.L. (VENTAS)', '', null, '', '', '')"
    ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 7.5, '', 91, 2, 'SILEMA BCN S.L. (MARKETING)', '', null, '', '', '')"
    ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 7.5, '', 91, 2, 'SILEMA BCN S.L. (MANTENIMIENTO)', '', null, '', '', '')"
    
    'FACTURA ---------------------------------------------------------------------------------------------------------------------
    FacturaClientSemanal client, data, data, data, Format(data, "dd-mm-yyyy"), "Live Preus Actuals", " Codiarticle in (" & articulo & ") and ", emisor
    

    '-------------------------------------------------------------------------------------------------------------------------------------------
    'CUOTA PROGRAMA TPV ------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------
    client = 1253 'SILEMA
    emisor = "19" 'SOLUCIONES
    articulo = "8332" 'Cuota Programa TPV
    
    'ALBARÁN ---------------------------------------------------------------------------------------------------------------------
    Set rsTiendas = Db.OpenResultset("select codi, nom from clients where codi in (select distinct botiga from (select * from [" & NomTaulaVentas(DateAdd("m", -1, Now())) & "] union all select * from [" & NomTaulaVentas(Now()) & "]) v) order by nom")
    While Not rsTiendas.EOF
        ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 1, '', 91, 2, 'SILEMA BCN S.L. (" & rsTiendas("nom") & ")[Lote:" & rsTiendas("nom") & "]', '', null, '', '', '')"
        rsTiendas.MoveNext
    Wend
    
    'FACTURA ---------------------------------------------------------------------------------------------------------------------
    FacturaClientSemanal client, data, data, data, Format(data, "dd-mm-yyyy"), "Live Preus Actuals", " Codiarticle in (" & articulo & ") and ", emisor


    '-------------------------------------------------------------------------------------------------------------------------------------------
    'CUOTA SOLUCIONES DIRECCIÓN ----------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------
    client = 1253 'SILEMA
    emisor = "19" 'SOLUCIONES
    articulo = "8333" 'Cuota Programa TPV
    
    'ALBARÁN ---------------------------------------------------------------------------------------------------------------------
    ExecutaComandaSql "insert into [" & DonamNomTaulaServit(data) & "] values (newid(), getdate(), '', " & client & ", " & articulo & ", null, 'SEMANAL', 'SOLUCIONES', 0, 0, 1, '', 91, 2, 'SILEMA BCN S.L.', '', null, '', '', '')"
    
    'FACTURA ---------------------------------------------------------------------------------------------------------------------
    FacturaClientSemanal client, data, data, data, Format(data, "dd-mm-yyyy"), "Live Preus Actuals", " Codiarticle in (" & articulo & ") and ", emisor

End Sub

Sub FacturaClientAsignaEmpresa()
   
'   ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   
   ExecutaComandaSql " Update TmpFactuacio Set empresa = null "


End Sub

   

Sub FacturaContador(An As Double, Optional empresa As Double = 0, Optional Cli As Double, Optional Ultima As Double, Optional UltimaNoAria As Double)
    Dim Rs As rdoResultset, Empresa2 As Double, Pre As String
    Dim client As Double
    
' Ultima = NumFacturaAdd(Empresa)
    client = Cli
    Empresa2 = empresa
    
    Set Rs = Db.OpenResultset("select * from constantsClient where codi=" & Cli & " and variable='SerieFacClient'")
    If Rs.EOF Then
        client = -2
    Else
        If Rs("valor") = "" Then client = -2
    End If
    
    FacturaContador2 An, Empresa2, client, Ultima, UltimaNoAria
    Pre = ""
    If Empresa2 <> 0 Then Pre = Empresa2 & "_"
    ExecutaComandaSql "Delete constantsempresa  where camp = '" & Pre & "CampSeguentFactura' "
    ExecutaComandaSql "insert into constantsempresa  (Camp,Valor) Values ('" & Pre & "CampSeguentFactura'," & Ultima & ") "
    ExecutaComandaSql "Delete constantsempresa  where camp = '" & Pre & "CampSeguentFacturaNoAria' "
    ExecutaComandaSql "insert into constantsempresa  (Camp,Valor) Values ('" & Pre & "CampSeguentFacturaNoAria'," & UltimaNoAria & ") "

End Sub


Sub FacturaContador2(An As Double, empresa As Double, Cli As Double, Optional Ultima As Double, Optional UltimaNoAria As Double)
'    Dim d As Date, i As Integer, rs As rdoResultset
    
'    Ultima = 0
'    UltimaNoAria = 0
'    For i = 1 To 12
'        d = DateSerial(An, i, 1)
'        If ExisteixTaula(NomTaulaFacturaData(d)) Then
'            If Cli <> -2 Then
'                Set rs = Db.OpenResultset("Select max(numfactura)  from [" & NomTaulaFacturaIva(d) & "] Where EmpresaCodi = " & Empresa & "  and ClientCodi = " & Cli & " ")
'            Else
'                Set rs = Db.OpenResultset("Select max(numfactura)  from [" & NomTaulaFacturaIva(d) & "] Where EmpresaCodi = " & Empresa & " ")
'            End If
'            If Not rs.EOF Then If Not IsNull(rs(0)) Then If rs(0) > Ultima Then Ultima = rs(0)
'
'            If Cli <> -2 Then
'                Set rs = Db.OpenResultset("Select Min(numfactura) from [" & NomTaulaFacturaIva(d) & "] Where EmpresaCodi = " & Empresa & "  and ClientCodi = " & Cli & "  ")
'            Else
'                Set rs = Db.OpenResultset("Select Min(numfactura) from [" & NomTaulaFacturaIva(d) & "] Where EmpresaCodi = " & Empresa & "  ")
'            End If
'            If Not rs.EOF Then If Not IsNull(rs(0)) Then If rs(0) < UltimaNoAria Then UltimaNoAria = rs(0)
'        End If
'    Next
'    Ultima = Ultima + 1
'    UltimaNoAria = UltimaNoAria - 1


    Dim D As Date, i As Integer, Rs As rdoResultset
    Dim serie As String, emp As String
    
    Ultima = 0
    UltimaNoAria = 0
    
    emp = ""
    serie = ""
    If empresa <> 0 Then emp = empresa & "_"
    Set Rs = Db.OpenResultset("select left(valor, 255) valor from ConstantsEmpresa where Camp = '" & emp & "CampSerieDeFactura'")
    If Not Rs.EOF Then serie = Rs("valor")
    
    For i = 1 To 12
        D = DateSerial(An, i, 1)
        If ExisteixTaula(NomTaulaFacturaData(D)) Then
            If Cli <> -2 Then
                Set Rs = Db.OpenResultset("Select max(numfactura)  from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & "  and ClientCodi = " & Cli & " ")
            Else
                If serie <> "" Then
                    Set Rs = Db.OpenResultset("Select max(numfactura)  from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " and serie='" & serie & "'")
                Else
                    Set Rs = Db.OpenResultset("Select max(numfactura)  from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " ")
                End If
            End If
            If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If Rs(0) > Ultima Then Ultima = Rs(0)
            
            If Cli <> -2 Then
                Set Rs = Db.OpenResultset("Select Min(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & "  and ClientCodi = " & Cli & "  ")
            Else
                If serie <> "" Then
                    Set Rs = Db.OpenResultset("Select Min(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " and serie='" & serie & "'")
                Else
                    Set Rs = Db.OpenResultset("Select Min(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & "  ")
                End If
            End If
            If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If Rs(0) < UltimaNoAria Then UltimaNoAria = Rs(0)
        End If
    Next
    Ultima = Ultima + 1
    UltimaNoAria = UltimaNoAria - 1

End Sub


Function CalculaVenciment(Forzat As String, DataFac As Date, client As Double, StDiesVenciment, StDiaPagament, FormaPagoLlista) As Date
   Dim DiesVenciment As Double, DiaPagament As Double
   
   CalculaVenciment = DataFac

   If Len(Forzat) > 0 And Forzat <> "0-0-" Then
      CalculaVenciment = DateSerial(Mid(Forzat, 7, 4), Mid(Forzat, 4, 2), Mid(Forzat, 1, 2))
   Else
      DiesVenciment = 0
      If Len(StDiesVenciment) > 0 Then If IsNumeric(StDiesVenciment) Then DiesVenciment = StDiesVenciment
      
      CalculaVenciment = DateAdd("d", DiesVenciment, CalculaVenciment)
      DiaPagament = Day(CalculaVenciment)
      If Len(StDiaPagament) > 0 Then If IsNumeric(StDiaPagament) Then DiaPagament = StDiaPagament
      
      While Day(CalculaVenciment) <> DiaPagament And DiaPagament > 0 And DiaPagament <= 31
         CalculaVenciment = DateAdd("d", 1, CalculaVenciment)
      Wend
      
   End If
    
    
End Function

Sub creaRebuts(data As Date, Idf As String, empresa As Double, client As Double)
    Dim sql As String, Rs As rdoResultset, Nrebuts As Double, Drebuts As Double, Domiciliat As Boolean, i As Integer
    Dim DataD As Date, rsCC As rdoResultset, cliCompte As String
    
    DataD = data
    Nrebuts = 1
    Set Rs = Db.OpenResultset("select Valor from constantsclient where codi = " & client & " and variable= 'Nrebuts' ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If IsNumeric(Rs(0)) Then Nrebuts = Rs(0)

    Drebuts = 1
    Set Rs = Db.OpenResultset("select Valor from constantsclient where codi = " & client & " and variable= 'Drebuts' ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If IsNumeric(Rs(0)) Then Drebuts = Rs(0)

   
    Domiciliat = True
    Set Rs = Db.OpenResultset("select Valor from constantsclient where codi = " & client & " and variable= 'FormaPagoLlista' ")
    If Not Rs.EOF Then
       If Not IsNull(Rs(0)) Then
          If Not Rs(0) = "" Then
             If Not Rs(0) = "1" Then
                Domiciliat = False
             End If
          End If
       End If
    End If
    
    sql = ""
    sql = sql & " delete [" & NomTaulaRebuts(DataD) & "] Where IdFactura = '" & Idf & "' "
    ExecutaComandaSql sql
   
    cliCompte = ""
    Set rsCC = Db.OpenResultset("select top 1 valor from constantsclient where variable = 'CompteCorrent' and codi = " & client)
    If Not rsCC.EOF Then cliCompte = rsCC("valor")
    If cliCompte = "" Then
        Set rsCC = Db.OpenResultset("select top 1 valor from constantsclient where variable = 'CompteCorrent' and cast(codi as nvarchar) in (select valor from constantsclient where variable = 'empMareFac' and codi=" & client & ")")
        If Not rsCC.EOF Then cliCompte = rsCC("valor")
    End If
    
   For i = 1 To Nrebuts
  
       sql = ""
       sql = sql & " Insert into [" & NomTaulaRebuts(DataD) & "]  "
       sql = sql & "([IdRebut],[DataCobrat],[Estat1],[Estat2],[Estat3],[Estat4],[Estat5],[IdFactura],[NumFactura],[EmpresaCodi],[Serie],[DataInici],[DataFi],[DataFactura],[DataEmissio],[DataVenciment],[FormaPagament],[Total],[ClientCodi],[ClientCodiFac],[ClientNom],[ClientNif],[ClientAdresa],[ClientCp],[Tel],[Fax],[eMail],[ClientLliure],[ClientCiutat],[ClientCompte],      [EmpNom],[EmpNif],[EmpAdresa],[EmpCp],[EmpTel],[EmpFax],[EmpeMail],[EmpLliure],[EmpCiutat],[CampMercantil],[EmpCompte],      [BaseIva1],[Iva1],[BaseIva2],[Iva2],[BaseIva3],[Iva3],[BaseIva4],[Iva4],      [BaseRec1],[Rec1],[BaseRec2],[Rec2],[BaseRec3],[Rec3],[BaseRec4],[Rec4],      [valorIva1],[valorIva2],[valorIva3],[valorIva4],      [valorRec1],[valorRec2],[valorRec3],[valorRec4],[IvaRec1],[IvaRec2],[IvaRec3],[IvaRec4],[Reservat]) "
       sql = sql & " Select newid(),null,'','','','','',[IdFactura],[NumFactura],[EmpresaCodi],[Serie],[DataInici],[DataFi],[DataFactura],[DataEmissio],[DataVenciment],[FormaPagament],[Total],[ClientCodi],[ClientCodiFac],[ClientNom],[ClientNif],[ClientAdresa],[ClientCp],[Tel],[Fax],[eMail],[ClientLliure],[ClientCiutat],'',      [EmpNom],[EmpNif],[EmpAdresa],[EmpCp],[EmpTel],[EmpFax],[EmpeMail],[EmpLliure],[EmpCiutat],[CampMercantil],'',      [BaseIva1],[Iva1],[BaseIva2],[Iva2],[BaseIva3],[Iva3],[BaseIva4],[Iva4],      [BaseRec1],[Rec1],[BaseRec2],[Rec2],[BaseRec3],[Rec3],[BaseRec4],[Rec4],      [valorIva1],[valorIva2],[valorIva3],[valorIva4],      [valorRec1],[valorRec2],[valorRec3],[valorRec4],[IvaRec1],[IvaRec2],[IvaRec3],[IvaRec4],      [Reservat]      from  [" & NomTaulaFacturaIva(data) & "] "
       sql = sql & " Where IdFactura = '" & Idf & "' "
       ExecutaComandaSql sql

       sql = ""
       sql = sql & " update [" & NomTaulaRebuts(DataD) & "] "
       sql = sql & " Set clientcompte = '" & cliCompte & "' "
       'sql = sql & " (select top 1 valor from constantsclient where variable = 'CompteCorrent' and codi = " & Client & ") "
       sql = sql & " Where IdFactura = '" & Idf & "' "
       ExecutaComandaSql sql

       sql = ""
       sql = sql & " update [" & NomTaulaRebuts(DataD) & "] "
       sql = sql & " Set EmpCompte = "
       sql = sql & " (select top 1 valor from constantsempresa Where camp = '"
       If empresa <> 0 Then sql = sql & empresa & "_"
       sql = sql & "CampCompteCorrent ') "
       sql = sql & " Where IdFactura = '" & Idf & "' "
       ExecutaComandaSql sql

       If Nrebuts > 1 Then
          If Nrebuts = i Then
            ExecutaComandaSql " update [" & NomTaulaRebuts(DataD) & "] Set Total = Total - (" & Nrebuts - 1 & " * round(Total / " & Nrebuts & ",2))  Where IdFactura = '" & Idf & "' "
          Else
            ExecutaComandaSql " update [" & NomTaulaRebuts(DataD) & "] Set Total = round(Total / " & Nrebuts & ",2) Where IdFactura = '" & Idf & "' "
          End If

       End If

       If Domiciliat = False Then ExecutaComandaSql " update [" & NomTaulaRebuts(DataD) & "] Set EmpCompte = '' Where IdFactura = '" & Idf & "' "
           DataD = DateAdd("d", Drebuts, DataD)
   Next



   
End Sub

Sub FacturaClientRecullDades(Cli As Double, Di As Date, Df As Date, Refacturar, PreusActuals, iD As String, MarcaClient As String, CondicioArticle As String, dataFact As Date)
    Dim D As Date, i As Integer, Tarifa, DescontePp, TipusFacturacio, sql As String, PosaViatge, familia, FamiliaDto As String
    Dim Rs As rdoResultset, article, ArticleDct, ArticleDto
     
    If UCase(EmpresaActual) = UCase("Pos3") Then RebaixaEstocksConjelador
     
    D = Di
    While D <= Df
        ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set Viatge = '' where viatge is null "
        
        PosaViatge = "' + Viatge + '"
        'If MarcaClient <> "" Then PosaViatge = MarcaClient
        
        If InStr(Refacturar, "CalRefacturar") > 0 Then
            sql = "Insert Into TmpFactuacio (IdFactura,Data,Producte,Client ,Servit,Tornat,Preu,Import,Referencia) "
            sql = sql & "Select '" & iD & "'," & sqlData(D) & " ,Codiarticle,Client,Sum(QuantitatServida) ,Sum(QuantitatTornada), "
            sql = sql & "case Left(comentari,6) when '[Tick:' then round(comentariper / case when Sum(QuantitatServida) =0 then 1 else Sum(QuantitatServida) end  ,3) else null end As Preu, "
            sql = sql & "case Left(comentari,6) when '[Tick:' then ComentariPer else null end As Import, "
            sql = sql & "left('[Centre:" & MarcaClient & "][Viatge:" & PosaViatge & "]' + case Left(comentari,6) when '[Tick:' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari when '[IdAlb' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari else '[Viatge:" & MarcaClient & "]' end + "
            sql = sql & "case when CHARINDEX('[Lote:', Comentari)>0 then "
            sql = sql & "SUBSTRING(comentari, charindex('[Lote:', comentari),CHARINDEX(']', comentari, charindex('[Lote:', comentari))-charindex('[Lote:', comentari)+1) else '' end ,255) as Referencia "
            sql = sql & "From [" & DonamNomTaulaServit(D) & "] "
            sql = sql & "Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) "
            'NO FACTURAR LAS REPOSICIONES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If UCase(EmpresaActual) = UCase("tena") Then
                sql = sql & "and (viatge <> 'Auto' or (viatge = 'Auto' and quantitatTornada <> 0)) "  'Solo los "Auto" que son devoluciones
            End If
            sql = sql & "Group by CodiArticle, Client, Comentari, Comentariper, viatge "
            ExecutaComandaSql sql
            
            ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set MotiuModificacio = '" & iD & "[TAB:" & NomTaulaFacturaIva(dataFact) & "]' Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) "  ' and isnull(MotiuModificacio,'') = '' "
        Else
            'ExecutaComandaSql "Insert Into TmpFactuacio (IdFactura,Data,Producte,Client ,Servit,Tornat,Preu,Import,Referencia) Select '" & id & "'," & SqlData(D) & " ,Codiarticle,Client,Sum(QuantitatServida) ,Sum(QuantitatTornada),case Left(comentari,6) when '[Tick:' then round(comentariper / case when Sum(QuantitatServida) =0 then 1 else Sum(QuantitatServida) end  ,3) else null end As Preu,case Left(comentari,6) when '[Tick:' then ComentariPer else null end As Import ,left('[Viatge:" & PosaViatge & "]' + case Left(comentari,6) when '[Tick:' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari when '[IdAlb' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari else '[Viatge:" & MarcaClient & "]' end,255) as REferencia  From [" & DonamNomTaulaServit(D) & "] Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) and isnull(MotiuModificacio,'') = '' Group by CodiArticle,Client,Comentari,Comentariper,viatge "
            'sql = "Insert Into TmpFactuacio (IdFactura,Data,Producte,Client ,Servit,Tornat,Preu,Import,Referencia) "
            'sql = sql & "Select '" & iD & "', " & sqlData(D) & " , Codiarticle, Client, Sum(QuantitatServida), Sum(QuantitatTornada), "
            'sql = sql & "case when charindex('[Preu:',comentari)>0 then substring(comentari,charIndex('[Preu:',comentari)+6 , (CHARINDEX(']',comentari,charindex('[Preu:', Comentari)+6))-(charIndex('[Preu:',comentari)+6))else case Left(comentari,6) when '[Tick:' then round(comentariper / case when Sum(QuantitatServida) =0 then 1 else Sum(QuantitatServida) end  ,3) else null end end As preu, "
            'sql = sql & "case Left(comentari,6) when '[Tick:' then ComentariPer else null end As Import, "
            'sql = sql & "left('[Centre:" & MarcaClient & "][Viatge:" & PosaViatge & "]' + case Left(comentari,6) when '[Tick:' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari when '[IdAlb' then '[Data:" & Format(D, "yyyy-mm-dd") & "]' + Comentari else '[Viatge:" & MarcaClient & "]' end + "
            'sql = sql & "case when CHARINDEX('[Lote:', Comentari)>0 then "
            'sql = sql & "SUBSTRING(comentari, charindex('[Lote:', comentari),CHARINDEX(']', comentari, charindex('[Lote:', comentari))-charindex('[Lote:', comentari)+1) else '' end ,255) as Referencia "
            'sql = sql & "From [" & DonamNomTaulaServit(D) & "] "
            'sql = sql & "Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) and len(isnull(MotiuModificacio,''))<3 Group by CodiArticle,Client,Comentari,Comentariper,viatge "
            
            sql = "Insert Into TmpFactuacio (IdFactura,Data,Producte,Client ,Servit,Tornat,Preu,Import,Referencia) "
            sql = sql & "Select '" & iD & "', " & sqlData(D) & " , Codiarticle, Client, Sum(QuantitatServida), Sum(QuantitatTornada), "
            sql = sql & "case when charindex('[Preu:',comentari)>0 then substring(comentari,charIndex('[Preu:',comentari)+6 , (CHARINDEX(']',comentari,charindex('[Preu:', Comentari)+6))-(charIndex('[Preu:',comentari)+6))else case Left(comentari,6) when '[Tick:' then round(comentariper / case when Sum(QuantitatServida) =0 then 1 else Sum(QuantitatServida) end  ,3) else null end end As preu, "
            sql = sql & "case Left(comentari,6) when '[Tick:' then ComentariPer else null end As Import, "
            sql = sql & "left('[Centre:" & MarcaClient & "][Viatge:" & PosaViatge & "][Data:" & Format(D, "yyyy-mm-dd") & "]' + comentari, 255) as Referencia "
            'sql = sql & "left('[Centre:" & MarcaClient & "][Viatge:" & PosaViatge & "]' + "
            'sql = sql & "case when CHARINDEX('[Tick:', Comentari)>0 then "
            'sql = sql & "'[Data:" & Format(D, "yyyy-mm-dd") & "]' + SUBSTRING(comentari, charindex('[Tick:', comentari),CHARINDEX(']', comentari, charindex('[Tick:', comentari))-charindex('[Tick:', comentari)+1) else '' end + "
            'sql = sql & "case when CHARINDEX('[IdAlbara:', Comentari)>0 then "
            'sql = sql & "'[Data:" & Format(D, "yyyy-mm-dd") & "]' + SUBSTRING(comentari, charindex('[IdAlbara:', comentari),CHARINDEX(']', comentari, charindex('[IdAlbara:', comentari))-charindex('[IdAlbara:', comentari)+1) else '' end + "
            'sql = sql & "case when CHARINDEX('[Lote:', Comentari)>0 then "
            'sql = sql & "SUBSTRING(comentari, charindex('[Lote:', comentari),CHARINDEX(']', comentari, charindex('[Lote:', comentari))-charindex('[Lote:', comentari)+1) else '' end + "
            'sql = sql & "case when CHARINDEX('[Repercutido:', Comentari)>0 then "
            'sql = sql & "SUBSTRING(comentari, charindex('[Repercutido:', comentari),CHARINDEX(']', comentari, charindex('[Repercutido:', comentari))-charindex('[Repercutido:', comentari)+1) else '' end "
            'sql = sql & ",255) as Referencia "
            sql = sql & "From [" & DonamNomTaulaServit(D) & "] "
            sql = sql & "Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) and len(isnull(MotiuModificacio,''))<3 "
            'NO FACTURAR LAS REPOSICIONES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If UCase(EmpresaActual) = UCase("tena") Then
                sql = sql & "and (viatge <> 'Auto' or (viatge = 'Auto' and quantitatTornada <> 0)) "  'Solo los "Auto" que son devoluciones
            End If
            sql = sql & "Group by CodiArticle, Client, Comentari, Comentariper, viatge "
                        
            ExecutaComandaSql sql
            
            ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set MotiuModificacio = '" & iD & "[TAB:" & NomTaulaFacturaIva(dataFact) & "]' Where " & CondicioArticle & " Client = " & Cli & " and (QuantitatServida<>0 or QuantitatTornada <> 0) and len(isnull(MotiuModificacio,''))<3 "
        End If
        D = DateAdd("d", 1, D)
        Debug.Print D
        DoEvents
    Wend
   
'   ExecutaComandaSql "Update TmpFactuacifo Set client = " & cli
    ExecutaComandaSql "Update TmpFactuacio Set Desconte = 0 Where client = " & Cli
   DoEvents
   For i = 0 To 100 Step 5
      ExecutaComandaSql "Update TmpFactuacio Set Desconte = " & i & " Where Referencia like '%Desc_" & i & "%' and client = " & Cli
      DoEvents
   Next
   
   ExecutaComandaSql "Update TmpFactuacio Set TipusIva = 2 , Acabat = 0 Where client = " & Cli
   DoEvents
   DoEvents   '
   ExecutaComandaSql "Update TmpFactuacio Set TipusIva = articles_zombis.TipoIva , Acabat = articles_zombis.NoDescontesEspecials From  TmpFactuacio Join articles_zombis on  TmpFactuacio.Producte = articles_zombis.Codi where TmpFactuacio.client = " & Cli
   DoEvents   '
   ExecutaComandaSql "Update TmpFactuacio Set TipusIva = Articles.TipoIva , Acabat = Articles.NoDescontesEspecials From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi  where TmpFactuacio.client = " & Cli
   DoEvents   '
   ExecutaComandaSql "Update TmpFactuacio Set Preu = Preu / ((100 + TI.Iva) / 100),Import = Import /  ((100 + TI.Iva) / 100) From  TmpFactuacio Join " & DonamTaulaTipusIva(dataFact) & " ti on  TmpFactuacio.TipusIva = TI.Tipus Where not Preu is null and not import is null  And client = " & Cli
   ExecutaComandaSql "Update TmpFactuacio Set Preu = Preu / ((100 - Desconte    ) / 100) From  TmpFactuacio Where not Preu is null and not import is null and Desconte > 0 and  client = " & Cli
   DoEvents
   ExecutaComandaSql "Update TmpFactuacio Set TipusIva = articles_zombis.TipoIva From  TmpFactuacio Join articles_zombis on  TmpFactuacio.Producte = articles_zombis.Codi Where client = " & Cli
   DoEvents
   ExecutaComandaSql "Update TmpFactuacio Set TipusIva = Articles.TipoIva From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where client = " & Cli
   DoEvents
   ExecutaComandaSql "Update Articles set Desconte = 1 where Desconte is null"
   Tarifa = ClientTarifaEspecial(Cli)
   DescontePp = ClientDescontePp(Cli)
   TipusFacturacio = ClientTipusFacturacio(Cli)
   
   For i = 1 To 4
      If Tarifa > 0 Then ExecutaComandaSql "Update TmpFactuacio Set " & ClientDesconteSql(Cli, i, "TarifesEspecials") & " From  TmpFactuacio Join TarifesEspecials On  TmpFactuacio.Producte = TarifesEspecials.Codi And TarifesEspecials.TarifaCodi = " & Tarifa & " Join Articles On TarifesEspecials.Codi = Articles.Codi And Articles.Desconte =  " & i & "  Where TmpFactuacio.client = " & Cli & " and TmpFactuacio.preu is null"
      DoEvents
      ExecutaComandaSql "Update TmpFactuacio Set " & ClientDesconteSql(Cli, i, "Articles_Zombis") & " from  TmpFactuacio Join Articles_Zombis on  TmpFactuacio.Producte = Articles_Zombis.Codi And Articles_Zombis.Desconte =  " & i & " Where TmpFactuacio.client = " & Cli & " and TmpFactuacio.preu is null"
      ExecutaComandaSql "Update TmpFactuacio Set " & ClientDesconteSql(Cli, i, "Articles") & " from  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi And Articles.Desconte =  " & i & " Where TmpFactuacio.client = " & Cli & " and TmpFactuacio.preu is null"
      DoEvents
   Next
   
   If Not PreusActuals Then
        For i = 1 To 4
            sql = ""
            sql = sql & "Update t set t.productenom = h2.nom,t.tipusiva=h2.tipoiva," & ClientDesconteSql(Cli, i, "h2") & " From  TmpFactuacio t join "
            sql = sql & "(select t.data,h.codi,min(h.fechamodif) modi from TmpFactuacio t join  articleshistorial h on t.producte = h.codi   and t.data <= h.fechamodif group by h.codi,t.data) a "
'            sql = sql & "(select t.data,h.codi,min(h.fechamodif) modi from TmpFactuacio t join  articleshistorial h on t.producte = h.codi  and h.Desconte = " & i & " where h.codi = t.producte and t.data < h.fechamodif group by h.codi,t.data ) a "
            sql = sql & " on t.producte = a.codi and t.data=a.data join articleshistorial h2 on h2.fechamodif = a.modi and h2.codi = a.codi  Where T.client = " & Cli
            
            If Tarifa > 0 Then sql = sql & " And Not t.producte in (select codi from TarifesEspecials where TarifaCodi = " & Tarifa & " )"
            ExecutaComandaSql sql
            DoEvents
        Next
        If Tarifa > 0 Then ExecutaComandaSql "Update  t set t.Preu = hh.PreuMajor From  TmpFactuacio t join ( select t.data,h.codi,min(h.fechamodif) modi from TmpFactuacio t join  TarifesHistorial h on t.producte = h.codi and h.TarifaCodi = " & Tarifa & " where h.codi = t.producte and t.data < h.fechamodif group by h.codi,t.data ) a on t.producte = a.codi and t.data=a.data join TarifesHistorial hh on hh.fechamodif = a.modi and hh.codi = a.codi  and hh.TarifaCodi = " & Tarifa & " Where t.Client = " & Cli
   End If
   
   
   
'******************** DESCUENTOS POR FAMILIAS
    Dim ii As Integer
    For ii = 1 To 3
        Set Rs = Db.OpenResultset("select valor as Descuento from constantsclient where variable = 'DtoFamilia' and codi=" & Cli)
        While Not Rs.EOF
            If Not IsNull(Rs("Descuento")) And Not Rs("Descuento") = "" Then
                familia = Split(Rs("Descuento"), "|")(0)
                FamiliaDto = Split(Rs("Descuento"), "|")(1)
                If FamiliaDto = "" Then FamiliaDto = 0
                If FamiliaDto >= 0 Then
                    If ii = 3 Then ExecutaComandaSql "Update TmpFactuacio Set Desconte = " & FamiliaDto & " From TmpFactuacio join articles On  TmpFactuacio.Producte = Articles.Codi And Articles.Familia = '" & familia & "' Where TmpFactuacio.Client = " & Cli
                    If ii = 2 Then ExecutaComandaSql "Update TmpFactuacio Set Desconte = " & FamiliaDto & " From TmpFactuacio Fac join articles A On  Fac.Producte = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom and F2.nom = '" & familia & "' Where Fac.Client = " & Cli
                    If ii = 1 Then ExecutaComandaSql "Update TmpFactuacio Set Desconte = " & FamiliaDto & " From TmpFactuacio Fac join articles A On  Fac.Producte = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom join families F1 on F2.pare = F1.nom and F1.nom = '" & familia & "' Where Fac.Client = " & Cli
                End If
            End If
            DoEvents
            Rs.MoveNext
        Wend
        Rs.Close
    Next
'******************** DESCUENTOS POR PRODUCTOS
    Set Rs = Db.OpenResultset("select valor as Descuento from constantsclient where variable = 'DtoProducte' and codi=" & Cli)
    While Not Rs.EOF
        If Not IsNull(Rs("Descuento")) Then
            article = Split(Rs("Descuento"), "|")(0)
            ArticleDto = Split(Rs("Descuento"), "|")(1)
'            If ArticleDto >= 0 Then
                ExecutaComandaSql "Update TmpFactuacio Set Desconte = " & ArticleDto & " From TmpFactuacio Where TmpFactuacio.Producte = '" & article & "' And TmpFactuacio.Client = " & Cli
'            End If
        End If
      DoEvents
      Rs.MoveNext
    Wend
    Rs.Close

    FiltraDevolucioMaxima Cli
    
' Fixem El Preu De La Tarifa Espècial
   sql = "Update TmpFactuacio set Preu = tarifesespecialsclients." & ClientTipusPreu(Cli) & " "
   sql = sql & "from TmpFactuacio join tarifesespecialsclients on TmpFactuacio.producte = tarifesespecialsclients.codi "
   sql = sql & "and tarifesespecialsclients.client = " & Cli & " "
   sql = sql & "where TmpFactuacio.client = " & Cli & " "
   ExecutaComandaSql sql

   ExecutaComandaSql "Update TmpFactuacio Set Tornat =  0 where Tornat is null "
   ExecutaComandaSql "Update TmpFactuacio Set Servit =  0 where Servit is null "
   ExecutaComandaSql "Delete TmpFactuacio where Servit = 0 And Tornat = 0 "
   ExecutaComandaSql "update f set  f.rec=0 , f.iva = i.iva from TmpFactuacio F join " & DonamTaulaTipusIva(dataFact) & " i on  f.tipusiva = i.tipus where f.client = " & Cli & ""
   If TipusFacturacio = 2 Then ExecutaComandaSql "update f set  f.rec = i.irpf from TmpFactuacio F join " & DonamTaulaTipusIva(dataFact) & " i on  f.tipusiva = i.tipus and Acabat = 0 where f.client = " & Cli
   ExecutaComandaSql "update f set  f.ProducteNom = '** ' + A.nom from TmpFactuacio F join articles_zombis a on  a.codi = f.producte where f.client = " & Cli
   ExecutaComandaSql "update f set  f.ProducteNom = A.nom from TmpFactuacio F join articles a on  a.codi = f.producte where f.client = " & Cli
   
    Dim agrupaAlbarans As Boolean
    agrupaAlbarans = False
    
    If DescontePp > 0 Then
        Set Rs = Db.OpenResultset("select ISNULL(valor, '') valor from ConstantsEmpresa where camp = 'AgrupaAlbaransDPP'")
        If Not Rs.EOF Then
            If Rs("valor") = "on" Then agrupaAlbarans = True
        End If
        Rs.Close
   
        If Not agrupaAlbarans Then
            ExecutaComandaSql "Insert Into TmpFactuacio_2 Select * From TmpFactuacio  Where client = " & Cli
            Dim compraExterna As Double
            Dim totalfactura As Double
            Dim nuevoPp As Double
            
            compraExterna = 0
            Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f left join articlesPropietats ap on f.Producte=ap.codiArticle and ap.variable='CompraExterna' where f.client='" & Cli & "' and ap.valor='on'")
            If Not Rs.EOF Then compraExterna = Rs("import")
            If compraExterna > 0 Then
                totalfactura = 0
                Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f where f.client='" & Cli & "'")
                If Not Rs.EOF Then totalfactura = Rs("import")
                nuevoPp = Round((((totalfactura * (1 - DescontePp / 100)) - compraExterna) / (totalfactura - compraExterna)) * 100, 2)
                If nuevoPp < 0 Then nuevoPp = 100
                If nuevoPp = 100 Then
                    ExecutaComandaSql "Update TmpFactuacio Set Tornat =  0,Servit = 0 From TmpFactuacio Where TmpFactuacio.client = " & Cli
                Else
                    ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,0),Servit =  round(Servit * (" & nuevoPp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 1 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
                    ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,2),Servit =  round(Servit * (" & nuevoPp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 0 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
                End If
            Else
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,0),Servit =  round(Servit * (100-" & DescontePp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 1 and TmpFactuacio.client = " & Cli
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,2),Servit =  round(Servit * (100-" & DescontePp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 0 and TmpFactuacio.client = " & Cli
            End If
            
            ExecutaComandaSql "Update t2 set t2.servit = t2.servit - t1.servit , t2.tornat = t2.tornat - t1.tornat from TmpFactuacio T1 Join TmpFactuacio_2 T2 on t2.idFactura = T1.idFactura and t1.data = t2.data and t1.client=t2.client and t1.producte = t2.producte and t1.iva = t2.iva and t1.rec = t2.rec And t1.preu = t2.preu and t1.desconte = t2.desconte and t1.referencia = t2.referencia where t2.client = " & Cli
            ExecutaComandaSql "delete TmpFactuacio_2 Where servit = 0  and  tornat = 0 "
       End If
   End If

   If TipusFacturacio = 1 Then ' iva inclos
        ExecutaComandaSql "update TmpFactuacio set preu = round(preu/(1.00 + (iva / 100)),3) "
   End If

   If TipusFacturacio = 4 Then ' Estranger
        ExecutaComandaSql "update TmpFactuacio set tipusIva=9, iva=0, rec=0 "
   End If

End Sub


Sub FacturaRecapitulativa(fIni As Date, fFin As Date, client As String)
    Dim empresa As String, empActual As String
    Dim aux As Date, f As Date
    Dim rsCli As rdoResultset, rsCliF As rdoResultset, rsTicks As rdoResultset, rsFR As rdoResultset
    Dim sql As String
    Dim codiCliF As String
    Dim numFactura As String
    Dim serie As String
    Dim rsId As rdoResultset
    Dim idFactura As String, idFacturaRE As String
    Dim m As Integer, D As Date, rsNumFact As rdoResultset, Ultima As Double, nFactura As Double, UltimaRE As Double, nFacturaRE As Double
    Dim baseIva1 As Double, iva1 As Double, baseIva2 As Double, iva2 As Double, baseIva3 As Double, iva3 As Double, Total As Double, sqlData As String
    Dim emailStr As String, emailCap As String, emailStrSE As String
    Dim tVenuts As String
    Dim mes As Integer
    Dim NumTick As String
    Dim rsCliRectificativa As rdoResultset, cliRectificativa As Integer, rsProdRectificativa As rdoResultset, prodRectificativa As String, rsNouCli As rdoResultset
    
    
    serie = "RC/"
    
    On Error GoTo err
    
    'Cliente y producto para la rectificativa
    Dim CliNomRe As String, CliNifRe As String, CliAdresaRe As String, CliCpRe As String, CliCiutatRe As String
    Set rsCliRectificativa = Db.OpenResultset("select codi, nom, isnull(nif, '') nif, isnull(Adresa, '') Adresa, isnull(cp, '') cp, isnull(ciutat, '') ciutat from clients where nom = 'CLIENTES CONTADO'")
    If rsCliRectificativa.EOF Then
        Set rsNouCli = Db.OpenResultset("Select Max(Codi) as maximo from (select codi from Clients union select codi from Clients_zombis) c")
        If Not rsNouCli.EOF Then
            cliRectificativa = rsNouCli("maximo")
            If cliRectificativa < 1000 Then cliRectificativa = 1000
            If cliRectificativa <> "" Then
                cliRectificativa = cliRectificativa + 1
            Else
                cliRectificativa = 1000
            End If
        Else
            cliRectificativa = 1000
        End If
        ExecutaComandaSql "INSERT INTO Clients (Codi,Nom,[tipus iva],[preu base]) Values (" & cliRectificativa & ",'CLIENTES CONTADO',3,2)"
        Set rsCliRectificativa = Db.OpenResultset("select codi, nom, isnull(nif, '') nif, isnull(Adresa, '') Adresa, isnull(cp, '') cp, isnull(ciutat, '') ciutat from clients where nom = 'CLIENTES CONTADO'")
    End If
    cliRectificativa = rsCliRectificativa("codi")
    CliNomRe = rsCliRectificativa("nom")
    CliNifRe = rsCliRectificativa("nif")
    CliAdresaRe = rsCliRectificativa("Adresa")
    CliCpRe = rsCliRectificativa("cp")
    CliCiutatRe = rsCliRectificativa("ciutat")
    
    
    prodRectificativa = "99999"
    
    Dim CliNom As String, CliNif As String, CliAdresa As String, CliCp As String, CliCiutat As String
    Set rsCli = Db.OpenResultset("select * from clients where codi=" & client)
    If Not rsCli.EOF Then
        CliNom = rsCli("nom")
        CliNif = rsCli("nif")
        CliAdresa = rsCli("Adresa")
        CliCp = rsCli("Cp")
        CliCiutat = rsCli("ciutat")
    End If
    
    Dim CliEmail As String
    Set rsCli = Db.OpenResultset("select * from constantsClient where codi = " & client & " and variable = 'eMail'")
    If Not rsCli.EOF Then CliEmail = rsCli("valor")
    
    Dim CliCodiContable As String
    CliCodiContable = client
    Set rsCli = Db.OpenResultset("select * from constantsClient where codi = " & client & " and variable = 'CodiContable'")
    If Not rsCli.EOF Then CliCodiContable = rsCli("valor")
    
    
    emailCap = "Facturas recapitulativas de  " & CliNom & "(" & Right("0" & Month(fIni), 2) & "/" & Year(fIni) & ")"
    
    Dim rsEmp As rdoResultset, rsEmpresaBD As rdoResultset
    Dim empNom As String, empNif As String, empAdresa As String, empCp As String, empTel As String, empFax As String
    Dim empEMail As String, empCiutat As String, CampMercantil As String, empPre As String, EmpSerie As String
    Dim strOtros As String, empresaBD As String
    
    strOtros = ""
    Set rsCliF = Db.OpenResultset("select isnull(valor, '') codiCliF from constantsclient where codi=" & client & " and variable='CFINAL' and isnull(valor, '')<>''")
    While Not rsCliF.EOF
        codiCliF = rsCliF("codiCliF")
        If strOtros <> "" Then strOtros = strOtros & " or "
        strOtros = strOtros & " otros like '%" & codiCliF & "%' "
        rsCliF.MoveNext
    Wend
    
    sql = "select Db from hit.dbo.web_empreses where nom ='" & EmpresaActual & "'"
    Set rsEmpresaBD = Db.OpenResultset(sql)
    If Not rsEmpresaBD.EOF Then empresaBD = rsEmpresaBD("db")
    
    If strOtros <> "" Then
        If Month(fIni) = Month(fFin) Then
            tVenuts = NomTaulaVentasBak(fIni, empresaBD)
        Else
            mes = 0
            tVenuts = "("
            For f = fIni To fFin
                If mes <> Month(f) Then
                    If f <> fIni Then tVenuts = tVenuts & " union all "
                    tVenuts = tVenuts & "select * from " & NomTaulaVentasBak(f, empresaBD) & " "
                    mes = Month(f)
                End If
            Next
            tVenuts = tVenuts & ")"
        End If
        sql = "select venut.data, venut.num_tick, c.codi botigaCodi, c.nom botiga, venut.import, (venut.Import)/(1 + t.iva/100) ImportSinIVA, ((venut.Import)/(1 + t.iva/100))/venut.quantitat preuUnitari, a.TipoIva, t.iva, venut.plu, a.nom producte, venut.quantitat, isnull(cc.valor, 999) emp "
        sql = sql & "from " & tVenuts & " venut "
        sql = sql & "left join clients c on venut.botiga = c.codi "
        sql = sql & "left join constantsclient cc on c.codi = cc.codi and variable = 'EmpresaVendes' "
        sql = sql & "left join articles a on venut.plu  = a.codi "
        sql = sql & "left join TipusIva2012 t on a.tipoIva = t.tipus "
        sql = sql & "left join " & TaulaTicksRecapitulats(fIni) & " tr on tr.dataTick=venut.data and tr.botiga=venut.botiga and tr.NumTick = venut.Num_tick "
        sql = sql & "where (" & strOtros & ") "
        sql = sql & "and venut.data between convert(datetime,'" & fIni & "', 103) "
        sql = sql & "and convert(datetime,'" & fFin & "',103) + convert(datetime,'23:59:59',8) "
        'NO SELECCIONAR LOS TICKETS YA FACTURADOS (TaulaTicksRecapitulats)
        sql = sql & "and tr.IdFactura is null "
        sql = sql & "order by isnull(cc.valor, 999), venut.data, venut.num_tick, c.nom"
        Set rsTicks = Db.OpenResultset(sql)
        
        empActual = ""
        emailStrSE = ""
        emailStr = ""
        NumTick = ""
        If Not rsTicks.EOF Then empresa = rsTicks("emp")
        While Not rsTicks.EOF
            If rsTicks("emp") = 999 Then
                If emailStrSE = "" Then emailStrSE = "<BR><TABLE BORDER='1'><TR><TD><B>BOTIGA</B></TD><TD><B>DATA</B></TD><TD><B>NUM_TICK</B></TD><TD><B>PRODUCTE</B></TD><TD><B>PREU</B></TD><TD><B>IVA</B></TD><TD><B>QUANTITAT</B></TD><TD><B>IMPORT</B></TD></TR>"
                emailStrSE = emailStrSE & "<TR><TD>" & rsTicks("botiga") & "</TD><TD>" & rsTicks("Data") & "</TD><TD>" & rsTicks("num_tick") & "</TD><TD>" & rsTicks("Producte") & "</TD><TD>" & rsTicks("preuUnitari") & "</TD><TD>" & rsTicks("iva") & "</TD><TD>" & rsTicks("Quantitat") & "</TD><TD>" & rsTicks("preuUnitari") * rsTicks("Quantitat") & "</TD></TR>"
            Else
                If rsTicks("emp") <> empActual Then
                    If empActual <> "" Then
                        baseIva1 = Round(baseIva1, 3): iva1 = Round(iva1, 3): baseIva2 = Round(baseIva2, 3): iva2 = Round(iva2, 3): baseIva3 = Round(baseIva3, 3): iva3 = Round(iva3, 3)
                        Total = baseIva1 + baseIva2 + baseIva3 + iva1 + iva2 + iva3
                        
                        sql = "update [" & NomTaulaFacturaIva(Now()) & "] set "
                        sql = sql & "Total = " & Total & ", BaseIva1 = " & baseIva1 & ", Iva1 = " & iva1 & ", BaseIva2 = " & baseIva2 & ", Iva2 = " & iva2 & ", BaseIva3 = " & baseIva3 & ", Iva3 = " & iva3 & ", BaseIva4 = 0, Iva4 = 0 , BaseRec1 = 0, "
                        sql = sql & "Rec1 = 0, BaseRec2 = 0, Rec2 = 0, BaseRec3 = 0, Rec3 = 0, BaseRec4 = 0, Rec4 = 0, valorIva1 = 4, valorIva2 = 10, valorIva3 = 21, valorIva4 = 0, valorRec1 = 0.5, valorRec2 = 1.4, "
                        sql = sql & "valorRec3 = 5.2, valorRec4 = 0, IvaRec1 = 0, IvaRec2 = 0, IvaRec3 = 0, IvaRec4 = 0 "
                        sql = sql & "where idFactura='" & idFactura & "'"
                        ExecutaComandaSql sql
                        
                        'FACTURA RECTIFICATIVA -------------------------------------------------------------------------
                        'Número de factura
                        UltimaRE = 0
                        For m = 1 To 12
                            D = DateSerial(Year(Now()), m, 1)
                            If ExisteixTaula(NomTaulaFacturaData(D)) Then
                                Set rsNumFact = Db.OpenResultset("Select max(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " and serie = '" & EmpSerie & "RE/'")
                                If Not rsNumFact.EOF Then If Not IsNull(rsNumFact(0)) Then If rsNumFact(0) > UltimaRE Then UltimaRE = rsNumFact(0)
                            End If
                        Next
                        nFacturaRE = UltimaRE + 1
                        
                        Set rsId = Db.OpenResultset("select newid() i")
                        If Not rsId.EOF Then idFacturaRE = rsId("i")
                        
                        sql = "insert into [" & NomTaulaFacturaIva(Now()) & "] (IdFactura, NumFactura, EmpresaCodi, Serie, DataInici, DataFi, DataFactura, DataEmissio, DataVenciment, FormaPagament, "
                        sql = sql & "Total, ClientCodi, ClientCodiFac, ClientNom, ClientNif, ClientAdresa, ClientCp, Tel, Fax, eMail, ClientLliure, ClientCiutat, EmpNom, EmpNif, EmpAdresa, EmpCp, "
                        sql = sql & "EmpTel, EmpFax, EmpeMail, EmpLliure, EmpCiutat, CampMercantil, BaseIva1, Iva1, BaseIva2, Iva2, BaseIva3, Iva3, BaseIva4, Iva4, BaseRec1, Rec1, BaseRec2, Rec2, "
                        sql = sql & "BaseRec3, Rec3, BaseRec4, Rec4, valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, IvaRec1, IvaRec2, IvaRec3, IvaRec4, "
                        sql = sql & "Reservat) values ('" & idFacturaRE & "', " & nFacturaRE & ", " & empActual & ", '" & EmpSerie & "RE/', getdate(), getdate(), getdate(), getdate(), getdate(), 'CONTADO', "
                        sql = sql & -1 * Total & ", " & cliRectificativa & ", " & cliRectificativa & ", '" & CliNomRe & "', '" & CliNifRe & "', '" & CliAdresaRe & "', '" & CliCpRe & "', '', '', '', '', '" & CliCiutatRe & "', '" & empNom & "', "
                        sql = sql & "'" & empNif & "', '" & empAdresa & "', '" & empCp & "', '" & empTel & "', '" & empFax & "', '" & empEMail & "', '', '" & empCiutat & "', '', "
                        sql = sql & -1 * baseIva1 & ", " & -1 * iva1 & ", " & -1 * baseIva2 & ", " & -1 * iva2 & ", " & -1 * baseIva3 & ", " & -1 * iva3 & ", 0, 0, "
                        sql = sql & "0, 0, 0, 0, 0, 0, 0, 0, 4, 10, 21, 10, 0.5, 1.4, 5.2, 0, 0, 0, 0, 0, '')"
                        ExecutaComandaSql sql
                        
                        sqlData = "insert into [" & NomTaulaFacturaData(Now()) & "] (idFactura , data, client, producte, ProducteNom, Acabat, preu, import, desconte, tipusIva, iva, rec, Referencia, servit, Tornat) values "
                        sqlData = sqlData & "('" & idFacturaRE & "', getdate(), " & cliRectificativa & ", " & prodRectificativa & ", 'Venta Personal. F " & nFactura & "', 0, 1, " & -1 * (baseIva1 + baseIva2 + baseIva3) & ", 0, 2, 10, 0, "
                        sqlData = sqlData & "'[Centre:" & CliNomRe & "][Viatge:" & CliNomRe & "]', "
                        sqlData = sqlData & "0, " & baseIva1 + baseIva2 + baseIva3 & ")"
                        ExecutaComandaSql sqlData
                        
                        
                        Set rsFR = Db.OpenResultset("select * from FacturacioComentaris where idFactura='" & idFacturaRE & "'")
                        If Not rsFR.EOF Then
                            ExecutaComandaSql "update FacturacioComentaris set comentari='[RECTIFICATIVA_DE:" & nFactura & "]' + comentari where idFactura='" & idFacturaRE & "'"
                        Else
                            ExecutaComandaSql "insert into FacturacioComentaris (idFactura, data, comentari, Cobrat) values ('" & idFacturaRE & "', getdate(), '[RECTIFICATIVA_DE:" & nFactura & "]', 'N')"
                        End If
                        
                        If emailStr = "" Then emailStr = "<BR><TABLE BORDER='1'><TR><TD><B>EMPRESA</B></TD><TD><B>FACTURA</B></TD><TD><B>TOTAL</B></TD>"
                        emailStr = emailStr & "<TR><TD>" & empNom & "</TD><TD>" & EmpSerie & serie & nFactura & "</TD><TD>" & Total & " &euro;</TD></TR>"
                        
                        EnviaFacturaEmail client, "", CStr(nFactura), "[" & NomTaulaFacturaIva(Now()) & "]", idFactura
                        
                        'ENVIAR A MURANO
                        ExecutaComandaSql "Insert Into FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFactura', 0, '[" & idFactura & "]', '[" & Now() & "]', '[" & nFactura & "]', '[" & NomTaulaFacturaIva(Now()) & "]', '')"
                        ExecutaComandaSql "Insert Into FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFactura', 0, '[" & idFacturaRE & "]', '[" & Now() & "]', '[" & nFacturaRE & "]', '[" & NomTaulaFacturaIva(Now()) & "]', '')"
                    End If
                
                    Set rsId = Db.OpenResultset("select newid() i")
                    If Not rsId.EOF Then idFactura = rsId("i")
                    ExecutaComandaSql "insert into [" & NomTaulaFacturaIva(Now()) & "] (IdFactura) values ('" & idFactura & "')"
            
                    empresa = rsTicks("emp")
                    
                    If empresa = "0" Then
                        empPre = ""
                    Else
                        empPre = empresa & "_"
                    End If
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampNom'")
                    If Not rsEmp.EOF Then empNom = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampNif'")
                    If Not rsEmp.EOF Then empNif = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampAdresa'")
                    If Not rsEmp.EOF Then empAdresa = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampCp'")
                    If Not rsEmp.EOF Then empCp = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampTel'")
                    If Not rsEmp.EOF Then empTel = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampFax'")
                    If Not rsEmp.EOF Then empFax = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampEmail'")
                    If Not rsEmp.EOF Then empEMail = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampCiutat'")
                    If Not rsEmp.EOF Then empCiutat = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampMercantil'")
                    If Not rsEmp.EOF Then CampMercantil = rsEmp("valor")
                    Set rsEmp = Db.OpenResultset("select isnull(valor, '') valor from constantsEmpresa where camp = '" & empPre & "CampSerieDeFactura'")
                    If Not rsEmp.EOF Then EmpSerie = rsEmp("valor")

                
                    'Número de factura
                    Ultima = 0
                    For m = 1 To 12
                        D = DateSerial(Year(Now()), m, 1)
                        If ExisteixTaula(NomTaulaFacturaData(D)) Then
                            Set rsNumFact = Db.OpenResultset("Select max(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " and serie = '" & EmpSerie & serie & "'")
                            If Not rsNumFact.EOF Then If Not IsNull(rsNumFact(0)) Then If rsNumFact(0) > Ultima Then Ultima = rsNumFact(0)
                        End If
                    Next
                    nFactura = Ultima + 1
            
                    sql = "update [" & NomTaulaFacturaIva(Now()) & "] set "
                    sql = sql & "numFactura=" & nFactura & ", "
                    sql = sql & "serie='" & EmpSerie & serie & "', "
                    sql = sql & "EmpresaCodi=" & empresa & ", "
                    sql = sql & "dataInici = '" & fIni & "', dataFi = '" & fFin & "', dataFactura=getdate(), dataEmissio=getdate(), dataVenciment = '" & fFin & "', "
                    sql = sql & "ClientCodi = " & client & ", clientcodifac = '" & CliCodiContable & "', tel = '', fax = '', email = '" & CliEmail & "', "
                    sql = sql & "ClientNom = '" & CliNom & "', ClientNif = '" & CliNif & "', ClientAdresa = '" & CliAdresa & "', ClientCp = '" & CliCp & "', ClientCiutat = '" & CliCiutat & "', "
                    sql = sql & "EmpNom = '" & empNom & "', EmpNif = '" & empNif & "', EmpAdresa = '" & empAdresa & "', EmpCp = '" & empCp & "', EmpTel = '" & empTel & "', EmpFax = '" & empFax & "', "
                    sql = sql & "empEMail = '" & empEMail & "' , empCiutat= '" & empCiutat & "', CampMercantil = '" & CampMercantil & "' "
                    sql = sql & "where idFactura='" & idFactura & "'"
                    ExecutaComandaSql sql
                                   
                    Total = 0: baseIva1 = 0: baseIva2 = 0: baseIva3 = 0: iva1 = 0: iva2 = 0: iva3 = 0
                    empActual = rsTicks("emp")
                End If
        
                sqlData = "insert into [" & NomTaulaFacturaData(Now()) & "] (idFactura , data, client, producte, ProducteNom, Acabat, preu, import, desconte, tipusIva, iva, rec, Referencia, servit, Tornat) values "
                sqlData = sqlData & "('" & idFactura & "', '" & rsTicks("Data") & "', " & client & ", " & rsTicks("plu") & ", '" & rsTicks("Producte") & "', 0, " & rsTicks("preuUnitari") & ", " & rsTicks("preuUnitari") * rsTicks("Quantitat") & ", 0, " & rsTicks("TipoIva") & ", " & rsTicks("iva") & ", 0, "
                sqlData = sqlData & "'[Data:" & Year(rsTicks("Data")) & "-" & Right("0" & Month(rsTicks("Data")), 2) & "-" & Right("0" & Day(rsTicks("Data")), 2) & "][IdAlbara:" & rsTicks("num_tick") & "]', "
                sqlData = sqlData & rsTicks("Quantitat") & ", 0)"
                ExecutaComandaSql sqlData
                    
                Select Case rsTicks("tipoIva")
                    Case 1: baseIva1 = baseIva1 + rsTicks("ImportSinIVA"): iva1 = iva1 + (rsTicks("ImportSinIVA") * rsTicks("iva") / 100)
                    Case 2: baseIva2 = baseIva2 + rsTicks("ImportSinIVA"): iva2 = iva2 + (rsTicks("ImportSinIVA") * rsTicks("iva") / 100)
                    Case 3: baseIva3 = baseIva3 + rsTicks("ImportSinIVA"): iva3 = iva3 + (rsTicks("ImportSinIVA") * rsTicks("iva") / 100)
                End Select
                    
                'Total = Total + rsTicks("import")
                
                'MARCAR EL TICKET COMO FACTURADO
                If rsTicks("num_tick") <> NumTick Then
                    ExecutaComandaSql "insert into " & TaulaTicksRecapitulats(rsTicks("Data")) & " (DataTick, Botiga, NumTick, IdFactura, DataFactura, Serie, NumFactura) values ('" & rsTicks("Data") & "', '" & rsTicks("botigaCodi") & "', '" & rsTicks("num_tick") & "', '" & idFactura & "', getdate(), '" & EmpSerie & serie & "', '" & nFactura & "')"
                End If
                
            End If
            NumTick = rsTicks("num_tick")
            
            rsTicks.MoveNext
        Wend
        
        If empActual <> "" And empActual <> "999" Then
            baseIva1 = Round(baseIva1, 3): iva1 = Round(iva1, 3): baseIva2 = Round(baseIva2, 3): iva2 = Round(iva2, 3): baseIva3 = Round(baseIva3, 3): iva3 = Round(iva3, 3)
            Total = baseIva1 + baseIva2 + baseIva3 + iva1 + iva2 + iva3
                        
            sql = "update [" & NomTaulaFacturaIva(Now()) & "] set "
            sql = sql & "Total = " & Total & ", BaseIva1 = " & baseIva1 & ", Iva1 = " & iva1 & ", BaseIva2 = " & baseIva2 & ", Iva2 = " & iva2 & ", BaseIva3 = " & baseIva3 & ", Iva3 = " & iva3 & ", BaseIva4 = 0, Iva4 = 0 , BaseRec1 = 0, "
            sql = sql & "Rec1 = 0, BaseRec2 = 0, Rec2 = 0, BaseRec3 = 0, Rec3 = 0, BaseRec4 = 0, Rec4 = 0, valorIva1 = 4, valorIva2 = 10, valorIva3 = 21, valorIva4 = 0, valorRec1 = 0.5, valorRec2 = 1.4, "
            sql = sql & "valorRec3 = 5.2, valorRec4 = 0, IvaRec1 = 0, IvaRec2 = 0, IvaRec3 = 0, IvaRec4 = 0 "
            sql = sql & "where idFactura='" & idFactura & "'"
            ExecutaComandaSql sql
            
            'FACTURA RECTIFICATIVA -------------------------------------------------------------------------------------------------------------
            'Número de factura
            UltimaRE = 0
            For m = 1 To 12
                D = DateSerial(Year(Now()), m, 1)
                If ExisteixTaula(NomTaulaFacturaData(D)) Then
                    Set rsNumFact = Db.OpenResultset("Select max(numfactura) from [" & NomTaulaFacturaIva(D) & "] Where EmpresaCodi = " & empresa & " and serie = '" & EmpSerie & "RE/'")
                    If Not rsNumFact.EOF Then If Not IsNull(rsNumFact(0)) Then If rsNumFact(0) > UltimaRE Then UltimaRE = rsNumFact(0)
                End If
            Next
            nFacturaRE = UltimaRE + 1
            
            Set rsId = Db.OpenResultset("select newid() i")
            If Not rsId.EOF Then idFacturaRE = rsId("i")
            
            sql = "insert into [" & NomTaulaFacturaIva(Now()) & "] (IdFactura, NumFactura, EmpresaCodi, Serie, DataInici, DataFi, DataFactura, DataEmissio, DataVenciment, FormaPagament, "
            sql = sql & "Total, ClientCodi, ClientCodiFac, ClientNom, ClientNif, ClientAdresa, ClientCp, Tel, Fax, eMail, ClientLliure, ClientCiutat, EmpNom, EmpNif, EmpAdresa, EmpCp, "
            sql = sql & "EmpTel, EmpFax, EmpeMail, EmpLliure, EmpCiutat, CampMercantil, BaseIva1, Iva1, BaseIva2, Iva2, BaseIva3, Iva3, BaseIva4, Iva4, BaseRec1, Rec1, BaseRec2, Rec2, "
            sql = sql & "BaseRec3, Rec3, BaseRec4, Rec4, valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, IvaRec1, IvaRec2, IvaRec3, IvaRec4, "
            sql = sql & "Reservat) values ('" & idFacturaRE & "', " & nFacturaRE & ", " & empActual & ", '" & EmpSerie & "RE/', getdate(), getdate(), getdate(), getdate(), getdate(), 'CONTADO', "
            sql = sql & -1 * Total & ", " & cliRectificativa & ", " & cliRectificativa & ", '" & CliNomRe & "', '" & CliNifRe & "', '" & CliAdresaRe & "', '" & CliCpRe & "', '', '', '', '', '" & CliCiutatRe & "', '" & empNom & "', "
            sql = sql & "'" & empNif & "', '" & empAdresa & "', '" & empCp & "', '" & empTel & "', '" & empFax & "', '" & empEMail & "', '', '" & empCiutat & "', '', "
            sql = sql & -1 * baseIva1 & ", " & -1 * iva1 & ", " & -1 * baseIva2 & ", " & -1 * iva2 & ", " & -1 * baseIva3 & ", " & -1 * iva3 & ", 0, 0, "
            sql = sql & "0, 0, 0, 0, 0, 0, 0, 0, 4, 10, 21, 10, 0.5, 1.4, 5.2, 0, 0, 0, 0, 0, '')"
            ExecutaComandaSql sql
            
            sqlData = "insert into [" & NomTaulaFacturaData(Now()) & "] (idFactura , data, client, producte, ProducteNom, Acabat, preu, import, desconte, tipusIva, iva, rec, Referencia, servit, Tornat) values "
            sqlData = sqlData & "('" & idFacturaRE & "', getdate(), " & cliRectificativa & ", " & prodRectificativa & ", 'Venta Personal. F " & nFactura & "', 0, 1, " & -1 * (baseIva1 + baseIva2 + baseIva3) & ", 0, 2, 10, 0, "
            sqlData = sqlData & "'[Centre:" & CliNomRe & "][Viatge:" & CliNomRe & "]', "
            sqlData = sqlData & "0, " & baseIva1 + baseIva2 + baseIva3 & ")"
            ExecutaComandaSql sqlData
            
            Set rsFR = Db.OpenResultset("select * from FacturacioComentaris where idFactura='" & idFacturaRE & "'")
            If Not rsFR.EOF Then
                ExecutaComandaSql "update FacturacioComentaris set comentari='[RECTIFICATIVA_DE:" & nFactura & "]' + comentari where idFactura='" & idFacturaRE & "'"
            Else
                ExecutaComandaSql "insert into FacturacioComentaris (idFactura, data, comentari, Cobrat) values ('" & idFacturaRE & "', getdate(), '[RECTIFICATIVA_DE:" & nFactura & "]', 'N')"
            End If

            
            If emailStr = "" Then emailStr = "<BR><TABLE border='1'><TR><TD><B>EMPRESA</B></TD><TD><B>FACTURA</B></TD><TD><B>TOTAL</B></TD>"
            emailStr = emailStr & "<TR><TD>" & empNom & "</TD><TD>" & EmpSerie & "RC/" & nFactura & "</TD><TD>" & Total & " &euro;</TD></TR>"
            
            EnviaFacturaEmail client, "", CStr(nFactura), "[" & NomTaulaFacturaIva(Now()) & "]", idFactura
            
           'ENVIAR A MURANO
            ExecutaComandaSql "Insert Into FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFactura', 0, '[" & idFactura & "]', '[" & Now() & "]', '[" & nFactura & "]', '[" & NomTaulaFacturaIva(Now()) & "]', '')"
            ExecutaComandaSql "Insert Into FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFactura', 0, '[" & idFacturaRE & "]', '[" & Now() & "]', '[" & nFacturaRE & "]', '[" & NomTaulaFacturaIva(Now()) & "]', '')"

        End If
            
        If emailStrSE <> "" Then emailStrSE = emailStrSE & "</TABLE>"
        
        sf_enviarMail "", "ana@solucionesit365.com", emailCap, "<H1>" & emailCap & "</H1>" & emailStr & emailStrSE, "", ""
        sf_enviarMail "", CliEmail, emailCap, "<H1>" & emailCap & "</H1>" & emailStr & emailStrSE, "", ""
    End If
    Exit Sub
    
err:
   sf_enviarMail "", "ana@solucionesit365.com", emailCap, "ERROR FacturaRecapitulativa" & err.Description, "", ""
End Sub

Sub FiltraDevolucioMaxima(Cli)
    Dim Rs As rdoResultset, familia, FamiliaDto, ii As Integer, article, ArticleDto
    Dim Afectats As String
    Dim Rs2
    Dim servit, Tornat
'******************** Pctje Devolucio Maxima Acceptada
    
    For ii = 1 To 3
        Set Rs = Db.OpenResultset("select valor as Descuento from constantsclient where variable = 'DevFamilia' and codi=" & Cli)
        While Not Rs.EOF
            If Not IsNull(Rs("Descuento")) And Rs("Descuento") <> "" Then
                familia = Split(Rs("Descuento"), "|")(0)
                FamiliaDto = Split(Rs("Descuento"), "|")(1)
                If IsNumeric(FamiliaDto) Then
                    If ii = 3 Then Set Rs2 = Db.OpenResultset("Select Distinct producte p,Sum(Tornat) t From TmpFactuacio join articles On  TmpFactuacio.Producte = Articles.Codi And Articles.Familia = '" & familia & "' Where TmpFactuacio.Client = " & Cli & "  group by producte  having sum(tornat) > (sum(servit) * " & CDbl(FamiliaDto) / 100 & ")")
                    If ii = 2 Then Set Rs2 = Db.OpenResultset("Select Distinct producte p,Sum(Tornat) t From TmpFactuacio Fac join articles A On  Fac.Producte = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom and F2.nom = '" & familia & "' Where Fac.Client = " & Cli & "  group by producte  having sum(tornat) > (sum(servit) * " & CDbl(FamiliaDto) / 100 & ")")
                    If ii = 1 Then Set Rs2 = Db.OpenResultset("Select Distinct producte p,Sum(Tornat) t From TmpFactuacio Fac join articles A On  Fac.Producte = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom join families F1 on F2.pare = F1.nom and F1.nom = '" & familia & "' Where Fac.Client = " & Cli & "  group by producte  having sum(tornat) > (sum(servit) * " & CDbl(FamiliaDto) / 100 & ")")
                    While Not Rs2.EOF
                        ExecutaComandaSql "Update TmpFactuacio  set tornat = 0 where producte = " & Rs2("p")
                        ExecutaComandaSql "Update f set Tornat = " & Round(Rs2("t") * (CDbl(FamiliaDto) / 100), 0) & " from (select top 1 * from TmpFactuacio where producte = " & Rs2("p") & ") f"
                        Rs2.MoveNext
                    Wend
                    Rs2.Close
                End If
            End If
            DoEvents
            Rs.MoveNext
        Wend
        Rs.Close
    Next

    Set Rs = Db.OpenResultset("select valor as Descuento from constantsclient where variable = 'DevProducte' and codi=" & Cli)
    While Not Rs.EOF
        If Not IsNull(Rs("Descuento")) Then
            article = Split(Rs("Descuento"), "|")(0)
            ArticleDto = Split(Rs("Descuento"), "|")(1)
            If IsNumeric(ArticleDto) Then
                Set Rs2 = Db.OpenResultset("Select Distinct producte p, Sum(Servit) s, Sum(Tornat) t From TmpFactuacio Where TmpFactuacio.Producte = '" & article & "' And TmpFactuacio.Client = " & Cli & " group by producte")
                If Not Rs2.EOF Then
                    servit = Rs2("s")
                    Tornat = Rs2("t")
                    If Tornat > Round(servit * (CDbl(ArticleDto) / 100), 0) Then
                        ExecutaComandaSql "Update TmpFactuacio  set tornat = 0 where producte = " & Rs2("p")
                        ExecutaComandaSql "Update f set Tornat = " & Round(Rs2("s") * (CDbl(ArticleDto) / 100), 0) & " from (select top 1 * from TmpFactuacio where producte = " & Rs2("p") & ") f"
                    End If
                End If
            End If
        End If
      DoEvents
      Rs.MoveNext
    Wend
    Rs.Close

    Set Rs = Db.OpenResultset("select Distinct producte p, Sum(Servit) s, Sum(Tornat) t from TmpFactuacio where TmpFactuacio.Producte in (select codiArticle from articlespropietats where variable = 'NoDevolucions' and valor = 'on') and TmpFactuacio.Client = " & Cli & " group by producte")
    While Not Rs.EOF
        ExecutaComandaSql "Update TmpFactuacio set tornat = 0 where producte = " & Rs("p") & " and TmpFactuacio.Client = " & Cli
        Rs.MoveNext
    Wend
    Rs.Close

End Sub

Function PreuAutomatic(c As Double) As Boolean
    Dim Rs As rdoResultset
    
    PreuAutomatic = False
    
    Set Rs = Db.OpenResultset("select * from Constantsclient where codi =  " & c & " and variable = 'PreuAutomatic'")
    
    If Not Rs.EOF Then If Not IsNull(Rs("Valor")) Then If Rs("Valor") = "pAuto" Then PreuAutomatic = True

End Function

Sub TipusDeIva(T_1 As Double, T_2 As Double, T_3 As Double, T_4 As Double, TR_1 As Double, TR_2 As Double, TR_3 As Double, TR_4 As Double, dataFact As Date)
   Dim Rs As rdoResultset
   
   T_1 = 4
   T_2 = 8
   T_3 = 18
   T_4 = 28

   TR_1 = 0.5
   TR_2 = 1
   TR_3 = 1.5
   TR_4 = 3.5

On Error GoTo nor
   Set Rs = Db.OpenResultset("Select * from " & DonamTaulaTipusIva(dataFact) & " order by tipus")
   While Not Rs.EOF
      Select Case (Rs("Tipus"))
         Case 1: T_1 = Rs("Iva"): TR_1 = Rs("Irpf")
         Case 2: T_2 = Rs("Iva"): TR_2 = Rs("Irpf")
         Case 3: T_3 = Rs("Iva"): TR_3 = Rs("Irpf")
         Case 4: T_4 = Rs("Iva"): TR_4 = Rs("Irpf")
      End Select
      Rs.MoveNext
   Wend
   Rs.Close
   'T_4 = 10  'IBEE
nor:
End Sub

Function DonamTaulaTipusIva(dia As Date) As String
    Dim sql As String
    Dim tablaTipusIvaNou As String
    Dim tablaTipusIvaAntic As String
    Dim tablaTipusIva2012 As String
    
    tablaTipusIvaNou = "TipusIva"
    tablaTipusIvaAntic = "TipusIvaAntic"
    tablaTipusIva2012 = "TipusIva2012"
    
    If Not ExisteixTaula(tablaTipusIvaAntic) Then
        sql = "CREATE TABLE [dbo].[" & tablaTipusIvaAntic & "] ("
        sql = sql & "[Tipus] [nvarchar] (255) NULL,"
        sql = sql & "[Iva] [float] NOT NULL ,"
        sql = sql & "[Irpf] [float] NOT NULL"
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql (sql)

        ExecutaComandaSql ("Insert Into " & tablaTipusIvaAntic & " (Tipus, Iva, Irpf) values('1',  4, 0.5)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaAntic & " (Tipus, Iva, Irpf) values('2',  7, 1)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaAntic & " (Tipus, Iva, Irpf) values('3', 16, 4)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaAntic & " (Tipus, Iva, Irpf) values('4',  0, 0)")
    End If
    
    If Not ExisteixTaula(tablaTipusIvaNou) Then
        sql = "CREATE TABLE [dbo].[" & tablaTipusIvaNou & "] ("
        sql = sql & "[Tipus] [nvarchar] (255) NULL,"
        sql = sql & "[Iva] [float] NOT NULL ,"
        sql = sql & "[Irpf] [float] NOT NULL"
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql (sql)

        ExecutaComandaSql ("Insert Into " & tablaTipusIvaNou & " (Tipus, Iva, Irpf) values('1',  4, 0.5)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaNou & " (Tipus, Iva, Irpf) values('2',  8, 1)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaNou & " (Tipus, Iva, Irpf) values('3', 18, 4)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIvaNou & " (Tipus, Iva, Irpf) values('4',  0, 0)")
    End If

    If Not ExisteixTaula(tablaTipusIva2012) Then
        sql = "CREATE TABLE [dbo].[" & tablaTipusIva2012 & "] ("
        sql = sql & "[Tipus] [nvarchar] (255) NULL,"
        sql = sql & "[Iva] [float] NOT NULL ,"
        sql = sql & "[Irpf] [float] NOT NULL"
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql (sql)

        ExecutaComandaSql ("Insert Into " & tablaTipusIva2012 & " (Tipus, Iva, Irpf) values('1', 4, 0.5)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIva2012 & " (Tipus, Iva, Irpf) values('2', 10, 1)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIva2012 & " (Tipus, Iva, Irpf) values('3', 21, 4)")
        ExecutaComandaSql ("Insert Into " & tablaTipusIva2012 & " (Tipus, Iva, Irpf) values('4',  0, 0)")
    End If

    If (Year(dia) < 2010) Or (Month(dia) < 7 And Year(dia) = 2010) Then
        DonamTaulaTipusIva = tablaTipusIvaAntic
    Else
        If (Year(dia) < 2012) Or (Month(dia) < 9 And Year(dia) = 2012) Then
            DonamTaulaTipusIva = tablaTipusIvaNou
        Else
            DonamTaulaTipusIva = tablaTipusIva2012
        End If
    End If
    
End Function


Sub CarregaDadesEmpresa(client, empCodi As Double, EmpSerie As String, empNom, empNif, empAdresa, empCp, EmpLliure, Tel, Fax, email, CampMercantil, empCiutat, Forsada)
   Dim Rs As rdoResultset
   Dim Prefixe As String
   
   'EmpCodi = 0
   EmpSerie = ""
   Prefixe = ""
   empNom = " "
   empNif = " "
   empAdresa = " "
   empCp = " "
   EmpLliure = " "
   Tel = " "
   Fax = " "
   email = " "
   CampMercantil = " "
   empCiutat = " "
   
   Set Rs = Db.OpenResultset("Select Camp from constantsempresa where camp like '%_CampCliente'  and valor = '" & client & "'")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Prefixe = Left(Rs(0), InStr(Rs(0), "_"))
   Rs.Close
   If Forsada > 0 Then Prefixe = Forsada & "_"
   If Len(Prefixe) > 0 Then empCodi = Left(Prefixe, Len(Prefixe) - 1)
   
   
   Set Rs = Db.OpenResultset("Select isnull(camp, '') camp, isnull(left(valor, 255), '') valor from ConstantsEmpresa order by camp")
   While Not Rs.EOF
      If Not IsNull(Rs(0)) And Not IsNull(Rs(1)) Then
         If Rs(0) = Prefixe & "CampNom" Then empNom = Rs(1)
         If Rs(0) = Prefixe & "CampNif" Then empNif = Rs(1)
         If Rs(0) = Prefixe & "CampAdresa" Then empAdresa = Rs(1)
         If Rs(0) = Prefixe & "CampCodiPostal" Then empCp = Rs(1)
         If Rs(0) = Prefixe & "CampLliure" Then EmpLliure = Rs(1)
         If Rs(0) = Prefixe & "CampTel" Then Tel = Rs(1)
         If Rs(0) = Prefixe & "CampFax" Then Fax = Rs(1)
         If Rs(0) = Prefixe & "CampMail" Then email = Rs(1)
         If Rs(0) = Prefixe & "CampMercantil" Then CampMercantil = Rs(1)
         If Rs(0) = Prefixe & "CampCiutat" Then empCiutat = Rs(1)
         If Rs(0) = Prefixe & "CampProvincia" Then empCiutat = empCiutat & ", " & Rs(1)
         If Rs(0) = Prefixe & "CampSerieDeFactura" Then EmpSerie = Rs(1)
      End If
      Rs.MoveNext
   Wend
   Rs.Close

End Sub

Sub CarregaDadesClient(Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClientFormaPagament, clientCiutat, ClientCodiFac, Devolucions, CliSerie)
    Dim Rs As rdoResultset, Rs2 As rdoResultset
    
    Set Rs = Db.OpenResultset("Select * from Clients Where Codi = " & Cli & " ")
    clientNom = " "
    ClientNif = " "
    clientAdresa = " "
    clientCp = " "
    ClientLliure = " "
    clientTel = " "
    clientFax = " "
    clienteMail = " "
    ClientFormaPagament = " "
    clientCiutat = " "
    ClientNomComercial = " "
    ClientCodiFac = Cli
    Devolucions = True
    CliSerie = ""
   
    If Not Rs.EOF Then
        If Not IsNull(Rs("Nom Llarg")) Then clientNom = Rs("Nom Llarg")
        If Not IsNull(Rs("Nom")) Then ClientNomComercial = Rs("Nom")
        
        If Not IsNull(Rs("Nif")) Then ClientNif = Rs("Nif")
        
        If UCase(EmpresaActual) = UCase("Hitrs") Then 'SI TIENE LAS FACTURAS AGRUPADAS COGEMOS LA DIRECCIÓN DEL CLIENTE "MADRE"
            Set Rs2 = Db.OpenResultset("select * from constantsclient where codi=" & Cli & " and variable='agruparFacturas' and valor='agruparFacturas'")
            If Not Rs2.EOF Then
                Set Rs2 = Db.OpenResultset("select * from constantsclient where codi=" & Cli & " and variable='empMareFac'")
                If Not Rs2.EOF Then
                    Set Rs2 = Db.OpenResultset("Select * from Clients Where Codi = " & Rs2("valor") & " ")
                    If Not Rs2.EOF Then
                        If Not IsNull(Rs2("Adresa")) Then clientAdresa = Rs2("Adresa")
                        If Not IsNull(Rs2("Cp")) Then clientCp = Rs2("Cp")
                        If Not IsNull(Rs2("Lliure")) Then ClientLliure = Rs2("Lliure")
                        If Not IsNull(Rs2("Ciutat")) Then clientCiutat = Rs2("Ciutat")
                    End If
                End If
            End If
            
            If clientAdresa = " " Then
                If Not IsNull(Rs("Adresa")) Then clientAdresa = Rs("Adresa")
                If Not IsNull(Rs("Cp")) Then clientCp = Rs("Cp")
                If Not IsNull(Rs("Lliure")) Then ClientLliure = Rs("Lliure")
                If Not IsNull(Rs("Ciutat")) Then clientCiutat = Rs("Ciutat")
            End If
        Else
            If Not IsNull(Rs("Adresa")) Then clientAdresa = Rs("Adresa")
            If Not IsNull(Rs("Cp")) Then clientCp = Rs("Cp")
            If Not IsNull(Rs("Lliure")) Then ClientLliure = Rs("Lliure")
            If Not IsNull(Rs("Ciutat")) Then clientCiutat = Rs("Ciutat")
        End If
    End If

    If ExisteixTaula("ConstantsClient") Then
        Set Rs = Db.OpenResultset("Select * from ConstantsClient Where Codi = " & Cli & "  ")
        While Not Rs.EOF
            If Not IsNull(Rs(0)) And Not IsNull(Rs(1)) Then
                If Rs("Variable") = "Tel" Then clientTel = Rs("Valor")
                If Rs("Variable") = "Fax" Then clientFax = Rs("Valor")
                If Rs("Variable") = "eMail" Then clienteMail = Rs("Valor")
                If Rs("Variable") = "FormaPago" Then ClientFormaPagament = Rs("Valor")
                If Rs("Variable") = "CodiContable" Then ClientCodiFac = Rs("Valor")
                If Rs("Variable") = "NoDevolucions" Then Devolucions = False
                If Rs("Variable") = "Provincia" Then clientCiutat = clientCiutat & ", " & Rs("Valor")
                If Rs("Variable") = "SerieFacClient" Then CliSerie = Rs("Valor")
            End If
            Rs.MoveNext
        Wend
        Rs.Close
    End If
   
End Sub


Sub CarregaDadesClientVenciment(Cli, DiesVenciment, DiaPagament, FormaPagoLlista)
   Dim Rs As rdoResultset
   
   If ExisteixTaula("ConstantsClient") Then
      Set Rs = Db.OpenResultset("Select * from ConstantsClient Where Codi = " & Cli & "  ")
      While Not Rs.EOF
         If Not IsNull(Rs(0)) And Not IsNull(Rs(1)) Then
            If Rs("Variable") = "Venciment" Then DiesVenciment = Rs("Valor")
            If Rs("Variable") = "DiaPagament" Then DiaPagament = Rs("Valor")
            If Rs("Variable") = "FormaPagoLlista" Then FormaPagoLlista = Rs("Valor")
         End If
         Rs.MoveNext
      Wend
      Rs.Close
   End If
   
End Sub



Sub CarregaDadesClientAgregats(Cli As Double, Clis() As String)
    Dim Rs As rdoResultset, cc() As String, nif As String, i As Integer
   
    ReDim Clis(0)
    Clis(0) = Cli
    nif = ""
    i = 1
    
    If UCase(EmpresaActual) = UCase("eurofleca") Then
      If ExisteixTaula("Clients") Then
          Set Rs = Db.OpenResultset("Select isnull(nif,'') nif from clients where codi = " & Cli)
          If Not Rs.EOF Then nif = Rs("nif")
          Rs.Close
    
          If nif <> "" Then
            Set Rs = Db.OpenResultset("Select codi from clients where nif = '" & nif & "' and codi <> " & Cli)
            While Not Rs.EOF
               ReDim Preserve Clis(i)
               Clis(i) = Rs("codi")
               i = i + 1
               Rs.MoveNext
            Wend
            Rs.Close
          End If
      End If
    Else
        If ExisteixTaula("Clients") Then
            Set Rs = Db.OpenResultset("Select isnull(nif,'') nif from clients where codi = " & Cli)
            If Not Rs.EOF Then nif = Rs("nif")
            Rs.Close
  
            If nif <> "" Then
                Set Rs = Db.OpenResultset("select c1.codi from constantsclient c1 left join constantsclient c2 on c1.codi=c2.codi and c2.variable='agruparFacturas' where c1.variable='empMareFac' and c1.valor=" & Cli & " and c2.valor='agruparFacturas' and c1.codi<>" & Cli)
                While Not Rs.EOF
                    ReDim Preserve Clis(i)
                    Clis(i) = Rs("codi")
                    i = i + 1
                    Rs.MoveNext
                Wend
                Rs.Close
            End If
        End If
    End If
'      If cli = 1088 Then Clis = Split(cli & ",1046,1045,1043,1050,1047,1048,1044,1055,1051,1054,1052,1049,1088", ",")
End Sub




Sub CalculaIvas(BaseIvaTipus_1, BaseIvaTipus_1_Iva, BaseIvaTipus_2, BaseIvaTipus_2_Iva, BaseIvaTipus_3, BaseIvaTipus_3_Iva, BaseIvaTipus_4, BaseIvaTipus_4_Iva, BaseRecTipus_1, BaseRecTipus_1_Rec, BaseRecTipus_2, BaseRecTipus_2_Rec, BaseRecTipus_3, BaseRecTipus_3_Rec, BaseRecTipus_4, BaseRecTipus_4_Rec, IvaRec1, IvaRec2, IvaRec3, IvaRec4, dataFact As Date, ibee As Boolean)
    Dim Rs As rdoResultset, T_1 As Double, T_2 As Double, T_3 As Double, T_4 As Double, TR_1 As Double, TR_2 As Double, TR_3 As Double, TR_4 As Double
   
    TipusDeIva T_1, T_2, T_3, T_4, TR_1, TR_2, TR_3, TR_4, dataFact
    'If IBEE Then
    '    T_4 = 10 'IBEE
    'Else
        Set Rs = Db.OpenResultset("Select * from " & DonamTaulaTipusIva(dataFact) & " where tipus=4")
        If Not Rs.EOF Then T_4 = Rs("Iva")
    'End If
   
    BaseIvaTipus_1_Iva = Round((BaseIvaTipus_1) * (T_1 / 100), 2)
    BaseIvaTipus_2_Iva = Round((BaseIvaTipus_2) * (T_2 / 100), 2)
    BaseIvaTipus_3_Iva = Round((BaseIvaTipus_3) * (T_3 / 100), 2)
    BaseIvaTipus_4_Iva = Round((BaseIvaTipus_4) * (T_4 / 100), 2)
    
    BaseRecTipus_1_Rec = Round(BaseRecTipus_1 * (TR_1 / 100), 2)
    BaseRecTipus_2_Rec = Round(BaseRecTipus_2 * (TR_2 / 100), 2)
    BaseRecTipus_3_Rec = Round(BaseRecTipus_3 * (TR_3 / 100), 2)
    BaseRecTipus_4_Rec = Round(BaseRecTipus_4 * (TR_4 / 100), 2)
    
    IvaRec1 = Round(BaseRecTipus_1 * (T_1 / 100), 2)
    IvaRec2 = Round(BaseRecTipus_2 * (T_2 / 100), 2)
    IvaRec3 = Round(BaseRecTipus_3 * (T_3 / 100), 2)
    IvaRec4 = Round(BaseRecTipus_4 * (T_4 / 100), 2)

End Sub

Function ClientTipusPreu(Cli As Double)
  Dim Rs As rdoResultset, Tip As Integer
  
  Tip = 0
  Set Rs = Db.OpenResultset("Select [Preu Base] From Clients Where Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Tip = Rs(0)
  Rs.Close
  
  ClientTipusPreu = "PREU"
  If Tip = 2 Then ClientTipusPreu = "PreuMajor"


'ClientTipusPreu = "PreuMajor"  ' Prova


End Function

Function ClientDesconteSql(Cli As Double, K As Integer, Taula As String)
  Dim Rs As rdoResultset, Tip, DescTe As Boolean, Camp As String
  Tip = 0
  
  Set Rs = Db.OpenResultset("Select [Desconte " & K & "] From Clients Where Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Tip = Rs(0)
  Rs.Close
  
  DescTe = False
  Set Rs = Db.OpenResultset("Select Variable From ConstantsClient Where Variable = 'descTE' and Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If Rs(0) = "descTE" Then DescTe = True
  Rs.Close
  
  ClientDesconteSql = " Preu = " & Taula & "." & ClientTipusPreu(Cli) & " "
  
  If DescTe Or (UCase(EmpresaActual) = UCase("Daunis")) Then
     ClientDesconteSql = ClientDesconteSql & " , Desconte = " & Tip & " "
  Else
     If Tip > 0 Then   ' Les tarifesespecials no porten desconte
        If Taula = "TarifesEspecials" Then
           ClientDesconteSql = ClientDesconteSql & " , Desconte = 0 "
        Else
          ClientDesconteSql = ClientDesconteSql & " , Desconte = " & Tip & " "
        End If
     End If
   End If
   
End Function



Function ClientTarifaEspecial(Cli As Double)
  Dim Rs As rdoResultset
  
  ClientTarifaEspecial = 0
  
  Set Rs = Db.OpenResultset("Select [Desconte 5] From Clients Where Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ClientTarifaEspecial = Rs(0)
  Rs.Close
  
End Function



Function ClientDescontePp(Cli As Double)
  Dim Rs As rdoResultset
  
  ClientDescontePp = 0
  
  Set Rs = Db.OpenResultset("Select [Desconte ProntoPago] From Clients Where Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ClientDescontePp = Rs(0)
  Rs.Close
  
End Function




Function ClientTipusFacturacio(Cli As Double)
  Dim Rs As rdoResultset

' 1 - Iva Inclos
' 2 - Amb recarreg
' 3 - Sense recarreg


  ClientTipusFacturacio = 2
  
  Set Rs = Db.OpenResultset("Select [Tipus Iva] From Clients Where Codi = " & Cli & " ")
  If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ClientTipusFacturacio = Rs(0)
  Rs.Close
  
End Function





Sub DesdeFins(Std As String, Di As Date, Df As Date)
   Dim Sd1 As String, Sd2 As String
   
   CalculaDesdeHasta Std, Sd1, Sd2
   Sd1 = Right(Trim(Sd1), Len(Trim(Sd1)) - 1)
   Sd2 = Left(Trim(Sd2), Len(Trim(Sd2)) - 1)
   
   Di = DateSerial(Mid(Sd1, 7, 4), Mid(Sd1, 4, 2), Mid(Sd1, 1, 2))
   Df = DateSerial(Mid(Sd2, 7, 4), Mid(Sd2, 4, 2), Mid(Sd2, 1, 2))
   
   
End Sub

Sub Factura(Dates As String, StClients As String, StDataFactura As String, StVenciment As String, Refacturar As String)
    Dim Di As Date, Df As Date, clients() As Double, dataFactura As Date, L As String, i As Integer, Venciment As Date, Rs
    Dim dppCliMare As Double
    Dim AlbaraUn As String
    Dim AlbaraUn2 As String
    Dim empCli
    
    DesdeFins Dates, Di, Df
    DesempaquetaLlista StClients, clients
    StDataFactura = Car(StDataFactura)
    StVenciment = Car(StVenciment)
    dataFactura = Now
    If Len(StDataFactura) > 0 Then dataFactura = DateSerial(Mid(StDataFactura, 7, 4), Mid(StDataFactura, 4, 2), Mid(StDataFactura, 1, 2))
'   If Len(StVenciment) > 0 Then Venciment = DateSerial(Mid(StVenciment, 7, 4), Mid(StVenciment, 4, 2), Mid(StVenciment, 1, 2))
   
    If InStr(Refacturar, "Albara[") > 0 Then
        AlbaraUn = Car(Right(Refacturar, Len(Refacturar) - InStr(Refacturar, "Albara") - 5))
        If Len(AlbaraUn) > 0 Then AlbaraUn = " comentari like '%IdAlbara:" & AlbaraUn & "%' "
    End If
    
    For i = 1 To UBound(clients)
        empCli = 0
        Set Rs = Db.OpenResultset("Select Camp from constantsempresa where camp like '%CampCliente'  and valor = '" & clients(i) & "'")
        If Not Rs.EOF Then
            If Not IsNull(Rs(0)) Then
                If InStr(Rs(0), "_") Then empCli = Left(Rs(0), InStr(Rs(0), "_") - 1)
            End If
        End If
        Rs.Close
    
        Set Rs = Db.OpenResultset("select distinct cast(Valor as nvarchar(255)) from articlespropietats where variable = 'EMP_FACTURA' and valor <> " & empCli)
        AlbaraUn2 = ""
        If Not AlbaraUn = "" Then AlbaraUn2 = AlbaraUn & " And "
        While Not Rs.EOF
            FacturaClient clients(i), Di, Df, dataFactura, StVenciment, Refacturar, " Codiarticle in (select distinct codiarticle from articlespropietats where variable = 'EMP_FACTURA' and valor =" & Rs(0) & ") And " & AlbaraUn2, Rs(0)
            Rs.MoveNext
        Wend
        Informa "Factura per " & BotigaCodiNom(clients(i))
        FacturaClient clients(i), Di, Df, dataFactura, StVenciment, Refacturar, AlbaraUn2, empCli
    Next
   
End Sub

Sub FacturaClientEmp(Cli As Double, Di As Date, Df As Date, DataDac As Date, Venciment As String, Refacturar As String, articles As String)
'NO SE USA ------------------------------------
   Dim BaseFactura2 As Double, BaseFactura As Double, BaseIvaTipus_1 As Double, BaseIvaTipus_2 As Double, BaseIvaTipus_3 As Double, BaseIvaTipus_4 As Double, BaseIvaTipus_1_Iva As Double, BaseIvaTipus_2_Iva As Double, BaseIvaTipus_3_Iva As Double, BaseIvaTipus_4_Iva As Double, iD As String, D As Date, i As Integer, Rs As rdoResultset, Q As rdoQuery, FacturaTotal As Double, BaseRecTipus_1 As Double, BaseRecTipus_2 As Double, BaseRecTipus_3 As Double, BaseRecTipus_4  As Double, BaseRecTipus_1_Rec As Double, BaseRecTipus_2_Rec As Double, BaseRecTipus_3_Rec As Double, BaseRecTipus_4_Rec As Double, NumFacNoAria As Double, NumFac As Double, DescontePp As Double, TipusFacturacio As Double, empCodi As Double, EmpSerie As String, Tot_BaseIvaTipus_1 As Double, Tot_BaseIvaTipus_2 As Double, Tot_BaseIvaTipus_3 As Double, Tot_BaseIvaTipus_4 As Double, Tot_BaseRecTipus_1 As Double, Tot_BaseRecTipus_2 As Double, Tot_BaseRecTipus_3  As Double, Tot_BaseRecTipus_4  As Double, VencimentActual As Date, AceptaDevolucions
   Dim clientNom As String, ClientNif As String, clientAdresa As String, clientCp As String, ClientLliure As String, empNom As String, empNif As String, empAdresa As String, empCp As String, EmpLliure As String, empTel, empFax, empEMail, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCampMercantil, empCiutat, ClientNomComercial As String, Tarifa As Integer, Impostos As Double, ClientCodiFact
   Dim sql As String, valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double, IvaRec1 As Double, IvaRec2 As Double, IvaRec3 As Double, IvaRec4 As Double, PreusActuals As Boolean, DiesVenciment, DiaPagament, FormaPagoLlista, Clis() As String
   Dim CliSerie As String
     
   FacturacioCreaTaulesBuides DataDac
   PreusActuals = False
   If InStr(Refacturar, "Preus Actuals") > 0 Then PreusActuals = True

'   CarregaDadesEmpresa cli, Articles, EmpCodi, EmpSerie, EmpNom, EmpNif, EmpAdresa, EmpCp, EmpLliure, EmpTel, EmpFax, EmpeMail, ClientCampMercantil, EmpCiutat, Articles
   CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
   CarregaDadesClientVenciment Cli, DiesVenciment, DiaPagament, FormaPagoLlista
   CarregaDadesClientAgregats Cli, Clis
   If clientNom = "" Then clientNom = ClientNomComercial
   If ClientNomComercial = "" Then ClientNomComercial = clientNom
   
   iD = Format(Now, "dd-mm-yy hh:mm:ss ")
   Set Rs = Db.OpenResultset("Select newid()")
   If Not Rs.EOF Then iD = Rs(0)
   Rs.Close
   
   ExecutaComandaSql "Update Articles set TipoIva = 2 where not (TipoIva = 1 or TipoIva = 2 or TipoIva = 3 or TipoIva = 4)"
   ExecutaComandaSql "Update Articles set desconte = 1 where not desconte in (1,2,3,4)"
   
   For i = 0 To UBound(Clis)
        CarregaDadesClient Clis(i), clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
        If clientNom = "" Then clientNom = ClientNomComercial
        If ClientNomComercial = "" Then ClientNomComercial = clientNom
        If UBound(Clis) = 0 Then clientNom = ""
        FacturaClientRecullDades Val(Clis(i)), Di, Df, Refacturar, PreusActuals, iD, ClientNomComercial, "", DataDac
   Next
   
   CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
   If Not AceptaDevolucions Then
        ExecutaComandaSql " Update TmpFactuacio Set Tornat = 0 "
        ExecutaComandaSql " Update TmpFactuacio_2 Set Tornat = 0 "
   End If
   
   If clientNom = "" Then clientNom = ClientNomComercial
   If ClientNomComercial = "" Then ClientNomComercial = clientNom
   Tarifa = ClientTarifaEspecial(Cli)
   DescontePp = ClientDescontePp(Cli)
   TipusFacturacio = ClientTipusFacturacio(Cli)
   
   FiltraDevolucioMaxima Cli
   
   ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   ExecutaComandaSql " Update TmpFactuacio Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   DoEvents

   BaseFactura = 0: BaseFactura2 = 0: BaseIvaTipus_1 = 0: BaseIvaTipus_2 = 0: BaseIvaTipus_3 = 0: BaseIvaTipus_4 = 0: FacturaTotal = 0: BaseRecTipus_1 = 0: BaseRecTipus_2 = 0: BaseRecTipus_3 = 0: BaseRecTipus_4 = 0: IvaRec1 = 0: IvaRec2 = 0: IvaRec3 = 0: IvaRec4 = 0
   Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura = Rs(0)
   Rs.Close
  
   Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio_2 ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura2 = Rs(0)
   Rs.Close
  
   If Abs(BaseFactura) > 0.0001 Or Abs(BaseFactura2) > 0.0001 Then
      If PreuAutomatic(Cli) Then   ' Si Preu Automatic, -> a negra preu 1
          ExecutaComandaSql "Update TmpFactuacio_2 Set Preu = Articles.Preu From TmpFactuacio_2 Join Articles on  TmpFactuacio_2.Producte = Articles.Codi "
          ' Fixem El Preu De La Tarifa Espècial  SEMPRE EL 1
          sql = "Update TmpFactuacio_2 set Preu = tarifesespecialsclients.PREU "
          sql = sql & "from TmpFactuacio_2 join tarifesespecialsclients on TmpFactuacio_2.producte = tarifesespecialsclients.codi "
          sql = sql & "and tarifesespecialsclients.client = " & Cli & " "
          ExecutaComandaSql sql
      End If
     
      Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Group By TipusIva")
      While Not Rs.EOF
         If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
         Select Case Rs("TipusIva")
            Case 1: BaseIvaTipus_1 = Rs("Ba")
            Case 2: BaseIvaTipus_2 = Rs("Ba")
            Case 3: BaseIvaTipus_3 = Rs("Ba")
            Case Else: BaseIvaTipus_4 = BaseIvaTipus_4 + Rs("Ba")
         End Select
         End If
         Rs.MoveNext
         DoEvents
      Wend
      Rs.Close
      
      If TipusFacturacio = 2 Then
        Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Where Acabat = 0 Group By TipusIva")
         While Not Rs.EOF
            If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
               Select Case Rs("TipusIva")
                  Case 1: BaseRecTipus_1 = Rs("Ba")
                  Case 2: BaseRecTipus_2 = Rs("Ba")
                  Case 3: BaseRecTipus_3 = Rs("Ba")
                  Case Else: BaseRecTipus_4 = BaseRecTipus_4 + Rs("Ba")
               End Select
            End If
            Rs.MoveNext
            DoEvents
         Wend
         Rs.Close
        BaseIvaTipus_1 = BaseIvaTipus_1 - BaseRecTipus_1
         BaseIvaTipus_2 = BaseIvaTipus_2 - BaseRecTipus_2
         BaseIvaTipus_3 = BaseIvaTipus_3 - BaseRecTipus_3
         BaseIvaTipus_4 = BaseIvaTipus_4 - BaseRecTipus_4
      End If
      
      CalculaIvas BaseIvaTipus_1, BaseIvaTipus_1_Iva, BaseIvaTipus_2, BaseIvaTipus_2_Iva, BaseIvaTipus_3, BaseIvaTipus_3_Iva, BaseIvaTipus_4, BaseIvaTipus_4_Iva, BaseRecTipus_1, BaseRecTipus_1_Rec, BaseRecTipus_2, BaseRecTipus_2_Rec, BaseRecTipus_3, BaseRecTipus_3_Rec, BaseRecTipus_4, BaseRecTipus_4_Rec, IvaRec1, IvaRec2, IvaRec3, IvaRec4, DataDac, False
      Impostos = BaseIvaTipus_1_Iva + BaseRecTipus_1_Rec + IvaRec1 + BaseIvaTipus_2_Iva + BaseRecTipus_2_Rec + IvaRec2 + BaseIvaTipus_3_Iva + BaseRecTipus_3_Rec + IvaRec3 + BaseIvaTipus_4_Iva + BaseRecTipus_4_Rec + IvaRec4
      
      If BaseFactura2 > 0.001 Then BaseFactura2 = BaseFactura2 + Impostos
      
      FacturaTotal = BaseIvaTipus_1 + BaseIvaTipus_2 + BaseIvaTipus_3 + BaseIvaTipus_4 _
                    + BaseRecTipus_1 + BaseRecTipus_2 + BaseRecTipus_3 + BaseRecTipus_4 _
                    + Impostos
      
      TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, DataDac
      
      NumFac = -1
      If Abs(BaseFactura) > 0.001 Then
         FacturaContador Year(DataDac), empCodi, NumFac
         VencimentActual = CalculaVenciment(Venciment, DataDac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)
'         ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(DataDac) & "] Where EmpresaCodi= " & EmpCodi & " and  [NumFactura] = " & NumFac
         Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataDac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
'[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4]
         Q.rdoParameters(0) = iD '[IdFactura]
         Q.rdoParameters(1) = empCodi '[EmpresaCodi]
         Q.rdoParameters(2) = EmpSerie '[Serie]
         Q.rdoParameters(3) = NumFac '[NumFactura]
         Q.rdoParameters(4) = Di '[DataInici]
         Q.rdoParameters(5) = Df '[DataFi]
         Q.rdoParameters(6) = DataDac '[DataFactura]
         Q.rdoParameters(7) = Now '[DataEmissio]
         Q.rdoParameters(8) = VencimentActual '[DataVenciment]
         Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
         Q.rdoParameters(10) = FacturaTotal '[Total]
         Q.rdoParameters(11) = Cli '[ClientCodi]
         Q.rdoParameters(12) = clientNom '[ClientNom]
         Q.rdoParameters(13) = ClientNif '[ClientNif]
         Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
         Q.rdoParameters(15) = clientCp '[ClientCp]
         Q.rdoParameters(16) = clientTel '[Tel]
         Q.rdoParameters(17) = clientFax '[Fax]
         Q.rdoParameters(18) = clienteMail '[email]
         Q.rdoParameters(19) = ClientLliure '[ClientLliure]
         Q.rdoParameters(20) = empNom '[EmpNom]
         Q.rdoParameters(21) = empNif '[EmpNif]
         Q.rdoParameters(22) = empAdresa '[EmpAdresa]
         Q.rdoParameters(23) = empCp '[EmpCp]
         Q.rdoParameters(24) = empTel '[EmpTel]
         Q.rdoParameters(25) = empFax '[EmpFax]
         Q.rdoParameters(26) = empEMail '[Empemail]
         Q.rdoParameters(27) = EmpLliure '[EmpLliure]
         Q.rdoParameters(28) = BaseIvaTipus_1 '[Base1]
         Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
         Q.rdoParameters(30) = BaseIvaTipus_2 '[Base2]
         Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
         Q.rdoParameters(32) = BaseIvaTipus_3 '[Base3]
         Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
         Q.rdoParameters(34) = BaseIvaTipus_4 '[Base4]
         Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
         Q.rdoParameters(36) = BaseRecTipus_1 '[Rec1]
         Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
         Q.rdoParameters(38) = BaseRecTipus_2 '[Rec2]
         Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
         Q.rdoParameters(40) = BaseRecTipus_3 '[Rec3]
         Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
         Q.rdoParameters(42) = BaseRecTipus_4 '[Rec4]
         Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
         Q.rdoParameters(44) = clientCiutat '[Rec4]
         Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
         Q.rdoParameters(46) = empCiutat '[Rec4]
         Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
         Q.rdoParameters(48) = valorIva1 '[valorIva1]
         Q.rdoParameters(49) = valorIva2 '[valorIva2]
         Q.rdoParameters(50) = valorIva3 '[valorIva3]
         Q.rdoParameters(51) = valorIva4 '[valorIva4]
         Q.rdoParameters(52) = valorRec1 '[valorRec1]
         Q.rdoParameters(53) = valorRec2 '[valorRec2]
         Q.rdoParameters(54) = valorRec3 '[valorRec3]
         Q.rdoParameters(55) = valorRec4 '[valorRec4]
         Q.rdoParameters(56) = IvaRec1 '[IvaRec1]
         Q.rdoParameters(57) = IvaRec2 '[IvaRec2]
         Q.rdoParameters(58) = IvaRec3 '[IvaRec3]
         Q.rdoParameters(59) = IvaRec4 '[IvaRec4]
         Q.rdoParameters(60) = "V1.20040304" '[Reservat]
         Q.Execute
         creaRebuts DataDac, iD, empCodi, Cli
      End If
      
      DoEvents
      
      If Abs(BaseFactura2) > 0.001 Then
          Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Group By TipusIva")
          While Not Rs.EOF
             If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                Select Case Rs("TipusIva")
                    Case 1: Tot_BaseIvaTipus_1 = Rs("Ba")
                    Case 2: Tot_BaseIvaTipus_2 = Rs("Ba")
                    Case 3: Tot_BaseIvaTipus_3 = Rs("Ba")
                    Case Else: Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_4 + Rs("Ba")
                End Select
             End If
            Rs.MoveNext
            DoEvents
        Wend
        Rs.Close
        If TipusFacturacio = 2 Then
           Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Where Acabat = 0 Group By TipusIva")
           While Not Rs.EOF
              If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                 Select Case Rs("TipusIva")
                    Case 1: Tot_BaseRecTipus_1 = Rs("Ba")
                    Case 2: Tot_BaseRecTipus_2 = Rs("Ba")
                    Case 3: Tot_BaseRecTipus_3 = Rs("Ba")
                    Case Else: Tot_BaseRecTipus_4 = Tot_BaseRecTipus_4 + Rs("Ba")
                 End Select
              End If
              Rs.MoveNext
              DoEvents
           Wend
           Rs.Close
           Tot_BaseIvaTipus_1 = Tot_BaseIvaTipus_1 - Tot_BaseRecTipus_1
           Tot_BaseIvaTipus_2 = Tot_BaseIvaTipus_2 - Tot_BaseRecTipus_2
           Tot_BaseIvaTipus_3 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_3
           Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_4
        End If
         
         FacturaContador Year(DataDac), empCodi, NumFac, NumFacNoAria
         VencimentActual = CalculaVenciment(Venciment, DataDac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)

         If Not NumFacNoAria = -1 Then ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(DataDac) & "] Where EmpresaCodi= " & empCodi & " and  [NumFactura] = " & NumFacNoAria
         Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataDac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
      
         
         Q.rdoParameters(0) = "Previsio_" + iD '[IdFactura]
         Q.rdoParameters(1) = empCodi '[EmpresaCodi]
         Q.rdoParameters(2) = EmpSerie '[Serie]
         Q.rdoParameters(3) = NumFacNoAria '[NumFactura]
         Q.rdoParameters(4) = Di '[DataInici]
         Q.rdoParameters(5) = Df '[DataFi]
         Q.rdoParameters(6) = DataDac '[DataFactura]
         Q.rdoParameters(7) = Now '[DataEmissio]
         Q.rdoParameters(8) = VencimentActual '[DataVenciment]
         Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
         Q.rdoParameters(10) = BaseFactura2 '[Total]
         Q.rdoParameters(11) = Cli '[ClientCodi]
         Q.rdoParameters(12) = clientNom '[ClientNom]
         Q.rdoParameters(13) = "" 'ClientNif '[ClientNif]
         Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
         Q.rdoParameters(15) = clientCp '[ClientCp]
         Q.rdoParameters(16) = clientTel '[Tel]
         Q.rdoParameters(17) = clientFax '[Fax]
         Q.rdoParameters(18) = clienteMail '[email]
         Q.rdoParameters(19) = ClientLliure '[ClientLliure]
         Q.rdoParameters(20) = empNom '[EmpNom]
         Q.rdoParameters(21) = "" 'EmpNif '[EmpNif]
         Q.rdoParameters(22) = empAdresa '[EmpAdresa]
         Q.rdoParameters(23) = empCp '[EmpCp]
         Q.rdoParameters(24) = empTel '[EmpTel]
         Q.rdoParameters(25) = empFax '[EmpFax]
         Q.rdoParameters(26) = empEMail '[Empemail]
         Q.rdoParameters(27) = EmpLliure '[EmpLliure]
         Q.rdoParameters(28) = Tot_BaseIvaTipus_1 'BaseIvaTipus_1 '[Base1]
         Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
         Q.rdoParameters(30) = Tot_BaseIvaTipus_2 'BaseIvaTipus_2 '[Base2]
         Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
         Q.rdoParameters(32) = Tot_BaseIvaTipus_3 'BaseIvaTipus_3 '[Base3]
         Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
         Q.rdoParameters(34) = Tot_BaseIvaTipus_4 'BaseIvaTipus_4 '[Base4]
         Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
         Q.rdoParameters(36) = Tot_BaseRecTipus_1 '[Rec1]
         Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
         Q.rdoParameters(38) = Tot_BaseRecTipus_2 '[Rec2]
         Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
         Q.rdoParameters(40) = Tot_BaseRecTipus_3 '[Rec3]
         Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
         Q.rdoParameters(42) = Tot_BaseRecTipus_4 '[Rec4]
         Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
         Q.rdoParameters(44) = clientCiutat '[Rec4]
         Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
         Q.rdoParameters(46) = empCiutat      '[Rec4]
         Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
         Q.rdoParameters(48) = valorIva1 '[valorIva1]
         Q.rdoParameters(49) = valorIva2 '[valorIva2]
         Q.rdoParameters(50) = valorIva3 '[valorIva3]
         Q.rdoParameters(51) = valorIva4 '[valorIva4]
         Q.rdoParameters(52) = valorRec1 '[valorRec1]
         Q.rdoParameters(53) = valorRec2     '[valorRec2]
         Q.rdoParameters(54) = valorRec3     '[valorRec3]
         Q.rdoParameters(55) = valorRec4     '[valorRec4]
         Q.rdoParameters(56) = IvaRec1       '[IvaRec1]
         Q.rdoParameters(57) = IvaRec2       '[IvaRec2]
         Q.rdoParameters(58) = IvaRec3       '[IvaRec3]
         Q.rdoParameters(59) = IvaRec4       '[IvaRec4]
         Q.rdoParameters(60) = "V1.20040305" '[Reservat]
   
         Q.Execute
         DoEvents
      
      End If
      
      ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataDac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio "
      
      DoEvents
      
      If DescontePp > 0 Then
         DoEvents
         ExecutaComandaSql "Update TmpFactuacio_2 Set IdFactura = 'Previsio_' + IdFactura  "
         DoEvents
         ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,3) " ' Where Import is null "
         DoEvents
         ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataDac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio_2 "
         DoEvents
      End If
      FacturaContador Year(DataDac), empCodi
   Else    ' Si no import desmarquem albarans ----------------------------------------
        D = Di
        While D <= Df
            ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set MotiuModificacio = '' Where Client = " & Cli & " And MotiuModificacio = '" & iD & "'  "
            D = DateAdd("d", 1, D)
            DoEvents
        Wend
   End If
   
   DoEvents

End Sub


Sub DelFactura(idFactura As String, sD As String)
    Dim Di As Date, Df As Date, Rs As rdoResultset, dia As Date, D As Date, i As Integer, EmpresaCodi As Double, clientCodi As Double
    Dim nFacturaOriginal As Long, numFactura As Integer, idFacturaR As String, sql As String, An As Integer, serie As String
    Dim rsFR As rdoResultset, rsMurano As rdoResultset, rsTicks As rdoResultset
    Dim facturaExportada As Boolean

    sD = Car(sD)
    dia = DateSerial(Mid(sD, 7, 2), Mid(sD, 4, 2), Mid(sD, 1, 2))
    Informa "Esborrant Factura " & idFactura

    facturaExportada = False
    If UCase(EmpresaActual) = UCase("Tena") Then
        Set rsMurano = Db.OpenResultset("select * from " & TaulaHistoricoMURANO(dia) & " where TipoExportacion='FACTURA' and param2='" & idFactura & "'")
        If Not rsMurano.EOF Then facturaExportada = True
    End If
    
    Set Rs = Db.OpenResultset("Select * From [" & NomTaulaFacturaIva(dia) & "] where idfactura ='" & idFactura & "'")
    Di = Now
    If Not Rs.EOF Then If Not IsNull(Rs("DataInici")) Then Di = Rs("DataInici")
    Df = Now
    If Not Rs.EOF Then If Not IsNull(Rs("DataFi")) Then Df = Rs("DataFi")
    EmpresaCodi = 0
    If Not Rs.EOF Then If Not IsNull(Rs("EmpresaCodi")) Then EmpresaCodi = Rs("EmpresaCodi")
    clientCodi = -2
    If Not Rs.EOF Then If Not IsNull(Rs("ClientCodi")) Then clientCodi = Rs("ClientCodi")
    If Not Rs.EOF Then If Not IsNull(Rs("Serie")) Then serie = Rs("Serie")
    If Not Rs.EOF Then If Not IsNull(Rs("numFactura")) Then nFacturaOriginal = Rs("numFactura")
    
    If facturaExportada Then  'NO SE PUEDE BORRAR. SE GENERA RECTIFICATIVA
        An = Year(dia)
   
        numFactura = 0
        For i = 12 To 1 Step -1
             D = DateSerial(An, i, 1)
             If ExisteixTaula(NomTaulaFacturaIva(D)) Then
                  Set rsFR = Db.OpenResultset("Select isnull(max(numfactura), 0) nFactura from [" & NomTaulaFacturaIva(D) & "] Where serie='" & serie & "RE/'")
                  If rsFR("nFactura") > numFactura Then
                     numFactura = rsFR("nFactura")
                 End If
             End If
        Next
        For i = 12 To 1 Step -1
             D = DateSerial(An + 1, i, 1)
             If ExisteixTaula(NomTaulaFacturaIva(D)) Then
                  Set rsFR = Db.OpenResultset("Select isnull(max(numfactura), 0) nFactura from [" & NomTaulaFacturaIva(D) & "] Where serie='" & serie & "RE/'")
                  If rsFR("nFactura") > numFactura Then
                     numFactura = rsFR("nFactura")
                 End If
             End If
        Next
        numFactura = numFactura + 1
    
        Set rsFR = Db.OpenResultset("select * from FacturacioComentaris where idFactura='" & idFactura & "'")
        If Not rsFR.EOF Then
            ExecutaComandaSql "update FacturacioComentaris set comentari='[RECTIFICATIVA:" & numFactura & "]' + comentari where idFactura='" & idFactura & "'"
        Else
            ExecutaComandaSql "insert into FacturacioComentaris (idFactura, data, comentari, Cobrat) values ('" & idFactura & "', getdate(), '[RECTIFICATIVA:" & numFactura & "]', 'N')"
        End If
         
        Set rsFR = Db.OpenResultset("select newid() as id")
        If Not rsFR.EOF Then idFacturaR = rsFR("id")
        
        sql = "Insert Into [" & NomTaulaFacturaIva(Now()) & "] "
        sql = sql & "select '" & idFacturaR & "', " & numFactura & ", empresacodi, serie+'RE/', dataInici, dataFi, getdate(), getdate(), dataVenciment, FormaPagament, (-1)*Total, clientCodi, ClientCodiFac, ClientNom, ClientNif, "
        sql = sql & "ClientAdresa, ClientCp, Tel, Fax, email, clientlliure, ClientCiutat, EmpNom, EmpNif, EmpAdresa, EmpCp, EmpTel, EmpFax, EmpeMail, EmpLLiure, EmpCiutat, CampMercantil, "
        sql = sql & "(-1)*BaseIva1, (-1)*Iva1, (-1)*BaseIva2, (-1)*Iva2, (-1)*BaseIva3, (-1)*Iva3, (-1)*BaseIva4, (-1)*Iva4, (-1)*BaseRec1, (-1)*Rec1, (-1)*BaseRec2, (-1)*Rec2, (-1)*BaseRec3, (-1)*Rec3, "
        sql = sql & "(-1)*BaseRec4, (-1)*Rec4, valorIva1, valoriva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, (-1)*IvaRec1, (-1)*IvaRec2, (-1)*IvaRec3, (-1)*IvaRec4, Reservat "
        sql = sql & "From [" & NomTaulaFacturaIva(dia) & "] "
        sql = sql & "where idfactura='" & idFactura & "'"
        ExecutaComandaSql sql
   
        sql = "Insert into [" & NomTaulaFacturaData(Now()) & "] (IdFactura, Data, Client, Producte, ProducteNom, Acabat, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) "
        sql = sql & "select '" & idFacturaR & "', getdate(), client, Producte, ProducteNom, Acabat, preu, (-1)*Import, Desconte, TipusIva, Iva, Rec, referencia, Tornat, Servit "
        sql = sql & "From [" & NomTaulaFacturaData(dia) & "] "
        sql = sql & "where idfactura='" & idFactura & "'"
        ExecutaComandaSql sql
        
        sql = "Insert into [" & NomTaulaFacturaReb(Now()) & "] "
        sql = sql & "select newid(), null, Estat1, Estat2, Estat3, Estat4, Estat5, '" & idFacturaR & "', " & numFactura & ", empresacodi, serie+'RE/', dataInici, dataFi, getdate(), getdate(), dataVenciment, FormaPagament, (-1)*Total, clientCodi, ClientCodiFac, ClientNom, ClientNif, "
        sql = sql & "ClientAdresa, ClientCp, Tel, Fax, email, clientlliure, ClientCiutat, ClientCompte, EmpNom, EmpNif, EmpAdresa, EmpCp, EmpTel, EmpFax, EmpeMail, EmpLLiure, EmpCiutat, CampMercantil, empCompte,"
        sql = sql & "(-1)*BaseIva1, (-1)*Iva1, (-1)*BaseIva2, (-1)*Iva2, (-1)*BaseIva3, (-1)*Iva3, (-1)*BaseIva4, (-1)*Iva4, (-1)*BaseRec1, (-1)*Rec1, (-1)*BaseRec2, (-1)*Rec2, (-1)*BaseRec3, (-1)*Rec3,"
        sql = sql & "(-1)*BaseRec4, (-1)*Rec4, valorIva1, valoriva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, (-1)*IvaRec1, (-1)*IvaRec2, (-1)*IvaRec3, (-1)*IvaRec4, Reservat "
        sql = sql & "from [" & NomTaulaFacturaReb(dia) & "] "
        sql = sql & "where idfactura='" & idFactura & "'"
        ExecutaComandaSql sql
        
        If Left(idFactura, 9) = "Previsio_" Then idFactura = Right(idFactura, Len(idFactura) - 9)
        
        Set rsFR = Db.OpenResultset("select * from FacturacioComentaris where idFactura='" & idFacturaR & "'")
        If Not rsFR.EOF Then
            ExecutaComandaSql "update FacturacioComentaris set comentari='[RECTIFICATIVA_DE:" & nFacturaOriginal & "]' + comentari where idFactura='" & idFacturaR & "'"
        Else
            ExecutaComandaSql "insert into FacturacioComentaris (idFactura, data, comentari, Cobrat) values ('" & idFacturaR & "', getdate(), '[RECTIFICATIVA_DE:" & nFacturaOriginal & "]', 'N')"
        End If
        
        
        While Di <= Df
           ExecutaComandaSql "Update [" & DonamNomTaulaServit(Di) & "] Set motiumodificacio = '' where motiumodificacio like '" & idFactura & "%' or motiumodificacio like 'Previsio_" & idFactura & "%' "
           Di = DateAdd("d", 1, Di)
           DoEvents
        Wend
   
        InsertFeineaAFer "SincroMURANOFactura", "[" & idFacturaR & "]", "[" & Now() & "]", "[" & numFactura & "]", "[" & NomTaulaFacturaIva(Now()) & "]"
    
    Else
       
       
       ExecutaComandaSql "Select * Into [" & NomTaulaFacturaIva(dia) & "_Bak]  From [" & NomTaulaFacturaIva(dia) & "]  where idfactura ='" & idFactura & "'"
       ExecutaComandaSql "Select * Into [" & NomTaulaFacturaData(dia) & "_Bak] From [" & NomTaulaFacturaData(dia) & "] where idfactura ='" & idFactura & "'"
       ExecutaComandaSql "Insert Into [" & NomTaulaFacturaIva(dia) & "_Bak]  Select * From [" & NomTaulaFacturaIva(dia) & "]  where idfactura ='" & idFactura & "'"
       ExecutaComandaSql "Insert into [" & NomTaulaFacturaData(dia) & "_Bak] Select * From [" & NomTaulaFacturaData(dia) & "] where idfactura ='" & idFactura & "'"
       
       ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(dia) & "]  where idfactura ='" & idFactura & "'"
       ExecutaComandaSql "Delete [" & NomTaulaFacturaData(dia) & "] where idfactura ='" & idFactura & "'"
       
       D = dia
       For i = 1 To 12
          ExecutaComandaSql "Delete [" & NomTaulaRebutData(D) & "] where idfactura ='" & idFactura & "'"
          D = DateAdd("m", 1, D)
       Next
       
       ExecutaComandaSql "Delete [FacturacioComentaris]             where idfactura ='" & idFactura & "'"
       
       If Left(idFactura, 9) = "Previsio_" Then idFactura = Right(idFactura, Len(idFactura) - 9)
       
       While Di <= Df
          ExecutaComandaSql "Update [" & DonamNomTaulaServit(Di) & "] Set motiumodificacio = '' where motiumodificacio like '" & idFactura & "%' or motiumodificacio like 'Previsio_" & idFactura & "%' "
          Di = DateAdd("d", 1, Di)

          DoEvents
       Wend
       
       FacturaContador Year(Df), EmpresaCodi, clientCodi
    
    End If
    
    'Si es factura recapitulativa, "liberar" los tiquets asociados
    ExecutaComandaSql "delete from " & TaulaTicksRecapitulats(dia) & " where idFactura='" & idFactura & "'"
    
End Sub



Sub FacturacioCreaTaulesBuides(D As Date)
   Dim NomTaula As String, sql As String
   
   NomTaula = NomTaulaFacturaData(D)
   If Not ExisteixTaula(NomTaula) Then
      sql = "CREATE TABLE [" & NomTaula & "] ( "
      sql = sql & " [IdFactura]    [nvarchar] (255) NULL , "
      sql = sql & " [Data]         [datetime] NULL ,"
      sql = sql & " [Client]       [float]    NULL , "
      sql = sql & " [Producte]     [float]    NULL , "
      sql = sql & " [ProducteNom]  [nvarchar] (255) NULL , "
      sql = sql & " [Acabat]       [float]    NULL , "
      sql = sql & " [Preu]         [float]    NULL , "
      sql = sql & " [Import]       [float]    NULL , "
      sql = sql & " [Desconte]     [float]    NULL , "
      sql = sql & " [TipusIva]     [float]    NULL , "
      sql = sql & " [Iva]          [float]    NULL , "
      sql = sql & " [rec]          [float]    NULL , "
      sql = sql & " [Referencia]   [nvarchar] (255) NULL , "
      sql = sql & " [Servit]       [float]    NULL , "
      sql = sql & " [Tornat]       [float]    NULL   "
      sql = sql & " ) ON [PRIMARY] "
      ExecutaComandaSql sql
   Else
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [ProducteNom]  [nvarchar] (255) NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [Iva]          [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [rec]          [float]    NULL "
   End If
   
   NomTaula = NomTaulaFacturaIva(D)
   If Not ExisteixTaula(NomTaula) Then
      sql = "CREATE TABLE [" & NomTaula & "] ( "
      sql = sql & " [IdFactura]    [nvarchar] (255) NULL , "
      sql = sql & " [NumFactura]   [float]    NULL , "
      sql = sql & " [EmpresaCodi]  [float]    NULL , "
      sql = sql & " [Serie]        [nvarchar] (255) NULL , "
      sql = sql & " [DataInici]    [datetime] NULL ,"
      sql = sql & " [DataFi]       [datetime] NULL ,"
      sql = sql & " [DataFactura]  [datetime] NULL ,"
      sql = sql & " [DataEmissio]  [datetime] NULL ,"
      sql = sql & " [DataVenciment] [datetime] NULL ,"
      sql = sql & " [FormaPagament] [nvarchar] (255) NULL , "
      sql = sql & " [Total]        [float]    NULL , "
     
      sql = sql & " [ClientCodi]   [float]    NULL , "
      sql = sql & " [ClientCodiFac][nvarchar] (255) NULL , "
      sql = sql & " [ClientNom]    [nvarchar] (255) NULL , "
      sql = sql & " [ClientNif]    [nvarchar] (255) NULL , "
      sql = sql & " [ClientAdresa] [nvarchar] (255) NULL , "
      sql = sql & " [ClientCp]     [nvarchar] (255) NULL , "
      sql = sql & " [Tel]          [nvarchar] (255) NULL , "
      sql = sql & " [Fax]          [nvarchar] (255) NULL , "
      sql = sql & " [eMail]        [nvarchar] (255) NULL , "
      sql = sql & " [ClientLliure] [nvarchar] (255) NULL , "
      sql = sql & " [ClientCiutat] [nvarchar] (255) NULL , "
      
      sql = sql & " [EmpNom]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpNif]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpAdresa]    [nvarchar] (255) NULL , "
      sql = sql & " [EmpCp]        [nvarchar] (255) NULL , "
      sql = sql & " [EmpTel]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpFax]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpeMail]     [nvarchar] (255) NULL , "
      sql = sql & " [EmpLliure]    [nvarchar] (255) NULL , "
      sql = sql & " [EmpCiutat]    [nvarchar] (255) NULL , "
      sql = sql & " [CampMercantil][nvarchar] (255) NULL , "
      
      sql = sql & " [BaseIva1]     [float]    NULL , "
      sql = sql & " [Iva1]         [float]    NULL , "
      sql = sql & " [BaseIva2]     [float]    NULL , "
      sql = sql & " [Iva2]         [float]    NULL , "
      sql = sql & " [BaseIva3]     [float]    NULL , "
      sql = sql & " [Iva3]         [float]    NULL , "
      sql = sql & " [BaseIva4]     [float]    NULL , "
      sql = sql & " [Iva4]         [float]    NULL , "
      
      sql = sql & " [BaseRec1]     [float]    NULL , "
      sql = sql & " [Rec1]         [float]    NULL , "
      sql = sql & " [BaseRec2]     [float]    NULL , "
      sql = sql & " [Rec2]         [float]    NULL , "
      sql = sql & " [BaseRec3]     [float]    NULL , "
      sql = sql & " [Rec3]         [float]    NULL , "
      sql = sql & " [BaseRec4]     [float]    NULL , "
      sql = sql & " [Rec4]         [float]    NULL , "
      
      sql = sql & " [valorIva1]         [float]    NULL ,  "
      sql = sql & " [valorIva2]         [float]    NULL ,  "
      sql = sql & " [valorIva3]         [float]    NULL ,  "
      sql = sql & " [valorIva4]         [float]    NULL ,  "
      
      sql = sql & " [valorRec1]         [float]    NULL ,  "
      sql = sql & " [valorRec2]         [float]    NULL ,  "
      sql = sql & " [valorRec3]         [float]    NULL ,  "
      sql = sql & " [valorRec4]         [float]    NULL ,  "
    
      sql = sql & " [IvaRec1]         [float]    NULL  , "
      sql = sql & " [IvaRec2]         [float]    NULL  , "
      sql = sql & " [IvaRec3]         [float]    NULL  , "
      sql = sql & " [IvaRec4]         [float]    NULL  , "
      
      sql = sql & " [Reservat]        [nvarchar] (255) NULL   "
      
      sql = sql & " ) ON [PRIMARY] "
      ExecutaComandaSql sql
   Else
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [EmpresaCodi]  [float]    NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [Serie]        [nvarchar] (255) NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [ClientCodiFac][nvarchar] (255) NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorIva1]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorIva2]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorIva3]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorIva4]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorRec1]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorRec2]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorRec3]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [valorRec4]       [float]    NULL "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [IvaRec1]         [float]    NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [IvaRec2]         [float]    NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [IvaRec3]         [float]    NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [IvaRec4]         [float]    NULL  "
      ExecutaComandaSql "Alter TABLE [" & NomTaula & "] Add [Reservat]        [nvarchar] (255) NULL "
   End If

   NomTaula = NomTaulaRebuts(D)
   If Not ExisteixTaula(NomTaula) Then
      sql = "CREATE TABLE [" & NomTaula & "] ( "
      sql = sql & " [IdRebut]       [nvarchar] (255) NULL , "
      sql = sql & " [DataCobrat]    [datetime]       NULL ,"
      sql = sql & " [Estat1]        [nvarchar] (255) NULL , "
      sql = sql & " [Estat2]        [nvarchar] (255) NULL , "
      sql = sql & " [Estat3]        [nvarchar] (255) NULL , "
      sql = sql & " [Estat4]        [nvarchar] (255) NULL , "
      sql = sql & " [Estat5]        [nvarchar] (255) NULL , "
      sql = sql & " [IdFactura]    [nvarchar] (255) NULL , "
      sql = sql & " [NumFactura]   [float]    NULL , "
      sql = sql & " [EmpresaCodi]  [float]    NULL , "
      sql = sql & " [Serie]        [nvarchar] (255) NULL , "
      sql = sql & " [DataInici]    [datetime] NULL ,"
      sql = sql & " [DataFi]       [datetime] NULL ,"
      sql = sql & " [DataFactura]  [datetime] NULL ,"
      sql = sql & " [DataEmissio]  [datetime] NULL ,"
      sql = sql & " [DataVenciment] [datetime] NULL ,"
      sql = sql & " [FormaPagament] [nvarchar] (255) NULL , "
      sql = sql & " [Total]        [float]    NULL , "
     
      sql = sql & " [ClientCodi]   [float]    NULL , "
      sql = sql & " [ClientCodiFac][nvarchar] (255) NULL , "
      sql = sql & " [ClientNom]    [nvarchar] (255) NULL , "
      sql = sql & " [ClientNif]    [nvarchar] (255) NULL , "
      sql = sql & " [ClientAdresa] [nvarchar] (255) NULL , "
      sql = sql & " [ClientCp]     [nvarchar] (255) NULL , "
      sql = sql & " [Tel]          [nvarchar] (255) NULL , "
      sql = sql & " [Fax]          [nvarchar] (255) NULL , "
      sql = sql & " [eMail]        [nvarchar] (255) NULL , "
      sql = sql & " [ClientLliure] [nvarchar] (255) NULL , "
      sql = sql & " [ClientCiutat] [nvarchar] (255) NULL , "
      sql = sql & " [ClientCompte] [nvarchar] (255) NULL , "
      
      sql = sql & " [EmpNom]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpNif]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpAdresa]    [nvarchar] (255) NULL , "
      sql = sql & " [EmpCp]        [nvarchar] (255) NULL , "
      sql = sql & " [EmpTel]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpFax]       [nvarchar] (255) NULL , "
      sql = sql & " [EmpeMail]     [nvarchar] (255) NULL , "
      sql = sql & " [EmpLliure]    [nvarchar] (255) NULL , "
      sql = sql & " [EmpCiutat]    [nvarchar] (255) NULL , "
      sql = sql & " [CampMercantil][nvarchar] (255) NULL , "
      sql = sql & " [EmpCompte]    [nvarchar] (255) NULL , "
      
      sql = sql & " [BaseIva1]     [float]    NULL , "
      sql = sql & " [Iva1]         [float]    NULL , "
      sql = sql & " [BaseIva2]     [float]    NULL , "
      sql = sql & " [Iva2]         [float]    NULL , "
      sql = sql & " [BaseIva3]     [float]    NULL , "
      sql = sql & " [Iva3]         [float]    NULL , "
      sql = sql & " [BaseIva4]     [float]    NULL , "
      sql = sql & " [Iva4]         [float]    NULL , "
      
      sql = sql & " [BaseRec1]     [float]    NULL , "
      sql = sql & " [Rec1]         [float]    NULL , "
      sql = sql & " [BaseRec2]     [float]    NULL , "
      sql = sql & " [Rec2]         [float]    NULL , "
      sql = sql & " [BaseRec3]     [float]    NULL , "
      sql = sql & " [Rec3]         [float]    NULL , "
      sql = sql & " [BaseRec4]     [float]    NULL , "
      sql = sql & " [Rec4]         [float]    NULL , "
      
      sql = sql & " [valorIva1]         [float]    NULL ,  "
      sql = sql & " [valorIva2]         [float]    NULL ,  "
      sql = sql & " [valorIva3]         [float]    NULL ,  "
      sql = sql & " [valorIva4]         [float]    NULL ,  "
      
      sql = sql & " [valorRec1]         [float]    NULL ,  "
      sql = sql & " [valorRec2]         [float]    NULL ,  "
      sql = sql & " [valorRec3]         [float]    NULL ,  "
      sql = sql & " [valorRec4]         [float]    NULL ,  "
    
      sql = sql & " [IvaRec1]         [float]    NULL  , "
      sql = sql & " [IvaRec2]         [float]    NULL  , "
      sql = sql & " [IvaRec3]         [float]    NULL  , "
      sql = sql & " [IvaRec4]         [float]    NULL  , "
      
      sql = sql & " [Reservat]        [nvarchar] (255) NULL   "
      
      sql = sql & " ) ON [PRIMARY] "
      ExecutaComandaSql sql
   End If

   ExecutaComandaSql "Drop Table TmpFactuacio "
   sql = "CREATE TABLE [TmpFactuacio] ( "
      sql = sql & " [IdFactura]    [nvarchar] (255) NULL , "
      sql = sql & " [Data]         [datetime] NULL ,"
      sql = sql & " [Client]       [float]    NULL , "
      sql = sql & " [Producte]     [float]    NULL , "
      sql = sql & " [Iva]          [float]    NULL , "
      sql = sql & " [rec]          [float]    NULL , "
      sql = sql & " [ProducteNom]  [nvarchar] (255) NULL , "
      sql = sql & " [Acabat]       [float]    NULL , "
      sql = sql & " [Preu]         [float]    NULL , "
      sql = sql & " [Import]       [float]    NULL , "
      sql = sql & " [Desconte]     [float]    NULL , "
      sql = sql & " [TipusIva]     [float]    NULL , "
      sql = sql & " [Referencia]   [nvarchar] (255) NULL , "
      sql = sql & " [Servit]       [float]    NULL , "
      sql = sql & " [Tornat]       [float]    NULL   "
      sql = sql & " ) ON [PRIMARY] "
   ExecutaComandaSql sql
   
   ExecutaComandaSql "Drop Table TmpFactuacio_2 "
   sql = "CREATE TABLE [TmpFactuacio_2] ( "
      sql = sql & " [IdFactura]    [nvarchar] (255) NULL , "
      sql = sql & " [Data]         [datetime] NULL ,"
      sql = sql & " [Client]       [float]    NULL , "
      sql = sql & " [Producte]     [float]    NULL , "
      sql = sql & " [Iva]          [float]    NULL , "
      sql = sql & " [rec]          [float]    NULL , "
      sql = sql & " [ProducteNom]  [nvarchar] (255) NULL , "
      sql = sql & " [Acabat]       [float]    NULL , "
      sql = sql & " [Preu]         [float]    NULL , "
      sql = sql & " [Import]       [float]    NULL , "
      sql = sql & " [Desconte]     [float]    NULL , "
      sql = sql & " [TipusIva]     [float]    NULL , "
      sql = sql & " [Referencia]   [nvarchar] (255) NULL , "
      sql = sql & " [Servit]       [float]    NULL , "
      sql = sql & " [Tornat]       [float]    NULL   "
      sql = sql & " ) ON [PRIMARY] "
   ExecutaComandaSql sql
   
      
End Sub



Sub FacturaClient(Cli As Double, Di As Date, Df As Date, DataDac As Date, Venciment As String, Refacturar As String, CondicioArticle As String, empresa)
   Dim BaseFactura2 As Double, BaseFactura As Double, BaseIvaTipus_1 As Double, BaseIvaTipus_2 As Double, BaseIvaTipus_3 As Double, BaseIvaTipus_4 As Double, BaseIvaTipus_1_Iva As Double, BaseIvaTipus_2_Iva As Double, BaseIvaTipus_3_Iva As Double, BaseIvaTipus_4_Iva As Double, iD As String, D As Date, i As Integer
   Dim Rs As rdoResultset, Q As rdoQuery, FacturaTotal As Double, BaseRecTipus_1 As Double, BaseRecTipus_2 As Double, BaseRecTipus_3 As Double, BaseRecTipus_4  As Double, BaseRecTipus_1_Rec As Double, BaseRecTipus_2_Rec As Double, BaseRecTipus_3_Rec As Double, BaseRecTipus_4_Rec As Double, NumFacNoAria As Double
   Dim NumFac As Double, DescontePp As Double, TipusFacturacio As Double, empCodi As Double, EmpSerie As String
   Dim CliSerie As String, Tot_BaseIvaTipus_1 As Double, Tot_BaseIvaTipus_2 As Double, Tot_BaseIvaTipus_3 As Double, Tot_BaseIvaTipus_4 As Double, Tot_BaseRecTipus_1 As Double, Tot_BaseRecTipus_2 As Double, Tot_BaseRecTipus_3  As Double, Tot_BaseRecTipus_4  As Double, VencimentActual As Date, AceptaDevolucions
   Dim clientNom As String, ClientNif As String, clientAdresa As String, clientCp As String, ClientLliure As String, empNom As String, empNif As String, empAdresa As String, empCp As String, EmpLliure As String, empTel, empFax, empEMail, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCampMercantil, empCiutat, ClientNomComercial As String, Tarifa As Integer, Impostos As Double, ClientCodiFact
   Dim sql As String, valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double, IvaRec1 As Double, IvaRec2 As Double, IvaRec3 As Double, IvaRec4 As Double, PreusActuals As Boolean, DiesVenciment, DiaPagament, FormaPagoLlista, Clis() As String
   Dim BaseIBEE As Double
      
   FacturacioCreaTaulesBuides DataDac
   PreusActuals = False
   If InStr(Refacturar, "Preus Actuals") > 0 Then PreusActuals = True
   
   CarregaDadesEmpresa Cli, empCodi, EmpSerie, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa
   CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
   CarregaDadesClientVenciment Cli, DiesVenciment, DiaPagament, FormaPagoLlista
   CarregaDadesClientAgregats Cli, Clis
   If clientNom = "" Then clientNom = ClientNomComercial
   If ClientNomComercial = "" Then ClientNomComercial = clientNom
   
   iD = Format(Now, "dd-mm-yy hh:mm:ss ")
   Set Rs = Db.OpenResultset("Select newid()")
   If Not Rs.EOF Then iD = Rs(0)
   Rs.Close
   
   ExecutaComandaSql "Update Articles set TipoIva = 2 where not (TipoIva = 1 or TipoIva = 2 or TipoIva = 3 or TipoIva = 4)"
   ExecutaComandaSql "Update Articles set desconte = 1 where not desconte in (1,2,3,4)"
   
   For i = 0 To UBound(Clis)
        CarregaDadesClient Clis(i), clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
        If clientNom = "" Then clientNom = ClientNomComercial
        If ClientNomComercial = "" Then ClientNomComercial = clientNom
        FacturaClientRecullDades Val(Clis(i)), Di, Df, Refacturar, PreusActuals, iD, ClientNomComercial, CondicioArticle, DataDac
        If UBound(Clis) = 0 Then clientNom = ""
   Next
   
   CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
   If Not AceptaDevolucions Then
        ExecutaComandaSql " Update TmpFactuacio Set Tornat = 0 "
        ExecutaComandaSql " Update TmpFactuacio_2 Set Tornat = 0 "
   End If
   
   'FiltraDevolucioMaxima Cli
   
   If clientNom = "" Then clientNom = ClientNomComercial
   If ClientNomComercial = "" Then ClientNomComercial = clientNom
   Tarifa = ClientTarifaEspecial(Cli)
   DescontePp = ClientDescontePp(Cli)
   
   'Repartir Blanco/Negro
   Dim agrupaAlbarans As Boolean
   agrupaAlbarans = False
   If DescontePp > 0 Then
        Set Rs = Db.OpenResultset("select ISNULL(valor, '') valor from ConstantsEmpresa where camp = 'AgrupaAlbaransDPP'")
        If Not Rs.EOF Then
            If Rs("valor") = "on" Then agrupaAlbarans = True
        End If
        Rs.Close
        
        If agrupaAlbarans Then
            sql = "Insert Into TmpFactuacio_2 "
            sql = sql & "select idfactura, max(data), '" & Cli & "', Producte, iva, rec, producteNom, acabat, preu, "
            sql = sql & "import, desconte, tipusiva, '', sum(Servit),sum(Tornat) "
            sql = sql & "From TmpFactuacio "
            sql = sql & "group by idfactura, Producte, iva, rec, producteNom, acabat, preu, import, desconte, tipusiva "
            ExecutaComandaSql sql
            ExecutaComandaSql "delete from TmpFactuacio "
            ExecutaComandaSql "insert into TmpFactuacio select * from TmpFactuacio_2"
 
            Dim compraExterna As Double
            Dim totalfactura As Double
            Dim nuevoPp As Double
                        
            compraExterna = 0
            Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f left join articlesPropietats ap on f.Producte=ap.codiArticle and ap.variable='CompraExterna' where f.client='" & Cli & "' and ap.valor='on'")
            If Not Rs.EOF Then compraExterna = Rs("import")
            If compraExterna > 0 Then
                totalfactura = 0
                Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f where f.client='" & Cli & "'")
                If Not Rs.EOF Then totalfactura = Rs("import")
                nuevoPp = Round((((totalfactura * (1 - DescontePp / 100)) - compraExterna) / (totalfactura - compraExterna)) * 100, 2)
                If nuevoPp < 0 Then nuevoPp = 100
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,0),Servit =  round(Servit * (" & nuevoPp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 1 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,2),Servit =  round(Servit * (" & nuevoPp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 0 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
            Else
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,0),Servit =  round(Servit * (100-" & DescontePp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 1 and TmpFactuacio.client = " & Cli
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,2),Servit =  round(Servit * (100-" & DescontePp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 0 and TmpFactuacio.client = " & Cli
            End If
   
            ExecutaComandaSql "Update t2 set t2.servit = t2.servit - t1.servit , t2.tornat = t2.tornat - t1.tornat from TmpFactuacio T1 Join TmpFactuacio_2 T2 on t2.idFactura = T1.idFactura and t1.data = t2.data and t1.client=t2.client and t1.producte = t2.producte and t1.iva = t2.iva and t1.rec = t2.rec And t1.preu = t2.preu and t1.desconte = t2.desconte and t1.referencia = t2.referencia where t2.client = " & Cli
            ExecutaComandaSql "delete TmpFactuacio_2 Where servit = 0  and  tornat = 0 "
        End If
    End If
   '~Repartir Blanco/Negro
    
   TipusFacturacio = ClientTipusFacturacio(Cli)
   
   If PreuAutomatic(Cli) Then   ' Si Preu Automatic, -> a negra preu 1
       ExecutaComandaSql "Update TmpFactuacio_2 Set Preu = Articles.Preu From TmpFactuacio_2 Join Articles on  TmpFactuacio_2.Producte = Articles.Codi "
       ' Fixem El Preu De La Tarifa Espècial  SEMPRE EL 1
       sql = "Update TmpFactuacio_2 set Preu = tarifesespecialsclients.PREU "
       sql = sql & "from TmpFactuacio_2 join tarifesespecialsclients on TmpFactuacio_2.producte = tarifesespecialsclients.codi "
       sql = sql & "and tarifesespecialsclients.client = " & Cli & " "
       ExecutaComandaSql sql
       If Tarifa > 0 Then ExecutaComandaSql "Update TmpFactuacio_2 Set Preu = TarifesEspecials.Preu From  TmpFactuacio_2 Join TarifesEspecials On  TmpFactuacio_2.Producte = TarifesEspecials.Codi And TarifesEspecials.TarifaCodi = " & Tarifa & " Join Articles On TarifesEspecials.Codi = Articles.Codi Where TmpFactuacio_2.client = " & Cli
   End If
   
   'Buscar productos con IBEE
   sql = "update TmpFactuacio_2 set referencia='[IBEE:' + isnull(ap.valor, '0') + ']' + referencia "
   sql = sql & "from TmpFactuacio_2 f "
   sql = sql & "left join ArticlesPropietats ap on f.Producte=ap.codiarticle and ap.variable='ENSUCRAT' "
   sql = sql & "where isnull(ap.valor, '') <> ''"
   ExecutaComandaSql sql

   sql = "update TmpFactuacio set referencia='[IBEE:' + isnull(ap.valor, '0') + ']' +referencia "
   sql = sql & "from TmpFactuacio f "
   sql = sql & "left join ArticlesPropietats ap on f.Producte=ap.codiarticle and ap.variable='ENSUCRAT' "
   sql = sql & "where isnull(ap.valor, '') <> ''"
   ExecutaComandaSql sql

   
   ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   ExecutaComandaSql " Update TmpFactuacio Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   
   'sql = "Update TmpFactuacio Set Import = Round(((f.Preu + isnull(ap.valor, 0) ) * (f.Servit - f.Tornat)) * (100-f.Desconte)/100,4) "
   'sql = sql & "from TmpFactuacio f "
   'sql = sql & "left join ArticlesPropietats ap on f.Producte=ap.codiarticle and ap.variable='ENSUCRAT' "
   'sql = sql & "where isnull(ap.valor, '') <> ''"
   'ExecutaComandaSql sql
   

   
   DoEvents
   
   BaseFactura = 0: BaseFactura2 = 0: BaseIvaTipus_1 = 0: BaseIvaTipus_2 = 0: BaseIvaTipus_3 = 0: BaseIvaTipus_4 = 0: FacturaTotal = 0: BaseRecTipus_1 = 0: BaseRecTipus_2 = 0: BaseRecTipus_3 = 0: BaseRecTipus_4 = 0: IvaRec1 = 0: IvaRec2 = 0: IvaRec3 = 0: IvaRec4 = 0
   BaseIBEE = 0
   Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura = Rs(0)
   Rs.Close
   
   Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio_2 ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura2 = Rs(0)
   Rs.Close
   
   If Abs(BaseFactura) > 0.0001 Or Abs(BaseFactura2) > 0.0001 Then
      Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Group By TipusIva")
      While Not Rs.EOF
         If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
            Select Case Rs("TipusIva")
                Case 1: BaseIvaTipus_1 = Rs("Ba")
                Case 2: BaseIvaTipus_2 = Rs("Ba")
                Case 3: BaseIvaTipus_3 = Rs("Ba")
                Case Else: BaseIvaTipus_4 = BaseIvaTipus_4 + Rs("Ba")
            End Select
         End If
         Rs.MoveNext
         DoEvents
      Wend
      Rs.Close
      
      Set Rs = Db.OpenResultset("select isnull(cast(replace(substring(referencia,charindex('IBEE:',referencia)+5,charindex(']',referencia,charindex('IBEE:',referencia)+5)-charindex('IBEE:',referencia)-5), ',','.') as float),0) * (Servit-Tornat) IBEE from TmpFactuacio where referencia like '%IBEE%'")
      While Not Rs.EOF
        BaseIBEE = BaseIBEE + Rs("IBEE")
        Rs.MoveNext
        DoEvents
      Wend
      Rs.Close
      
      If TipusFacturacio = 2 Then
         Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Where Acabat = 0 Group By TipusIva")
         While Not Rs.EOF
            If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
               Select Case Rs("TipusIva")
                  Case 1: BaseRecTipus_1 = Rs("Ba")
                  Case 2: BaseRecTipus_2 = Rs("Ba")
                  Case 3: BaseRecTipus_3 = Rs("Ba")
                  Case Else: BaseRecTipus_4 = BaseRecTipus_4 + Rs("Ba")
               End Select
            End If
            Rs.MoveNext
            DoEvents
         Wend
         Rs.Close
         BaseIvaTipus_1 = BaseIvaTipus_1 - BaseRecTipus_1
         BaseIvaTipus_2 = BaseIvaTipus_2 - BaseRecTipus_2
         BaseIvaTipus_3 = BaseIvaTipus_3 - BaseRecTipus_3
         BaseIvaTipus_4 = BaseIvaTipus_4 - BaseRecTipus_4
      End If
      
      CalculaIvas BaseIvaTipus_1, BaseIvaTipus_1_Iva, BaseIvaTipus_2, BaseIvaTipus_2_Iva, BaseIvaTipus_3, BaseIvaTipus_3_Iva, BaseIvaTipus_4, BaseIvaTipus_4_Iva, BaseRecTipus_1, BaseRecTipus_1_Rec, BaseRecTipus_2, BaseRecTipus_2_Rec, BaseRecTipus_3, BaseRecTipus_3_Rec, BaseRecTipus_4, BaseRecTipus_4_Rec, IvaRec1, IvaRec2, IvaRec3, IvaRec4, DataDac, False
      Impostos = BaseIvaTipus_1_Iva + BaseRecTipus_1_Rec + IvaRec1 + BaseIvaTipus_2_Iva + BaseRecTipus_2_Rec + IvaRec2 + BaseIvaTipus_3_Iva + BaseRecTipus_3_Rec + IvaRec3 + BaseIvaTipus_4_Iva + BaseRecTipus_4_Rec + IvaRec4
      
      If Not (UCase(EmpresaActual) = UCase("Tena") Or Year(Di) < 2010) Then
      
      Else
          If BaseFactura2 > 0.001 Then BaseFactura2 = BaseFactura2 + Impostos
      End If

      
      FacturaTotal = BaseIvaTipus_1 + BaseIvaTipus_2 + BaseIvaTipus_3 + BaseIvaTipus_4 _
                    + BaseRecTipus_1 + BaseRecTipus_2 + BaseRecTipus_3 + BaseRecTipus_4 _
                    + Impostos + BaseIBEE + (BaseIBEE * 0.1)
      
      TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, DataDac
      
      If TipusFacturacio = 4 Then
         BaseIvaTipus_1 = 0
         BaseIvaTipus_2 = 0
         BaseIvaTipus_3 = 0
         BaseIvaTipus_4 = 0
      End If
      
      NumFac = -1
      If Abs(BaseFactura) > 0.001 Then
         FacturaContador Year(DataDac), empCodi, Cli, NumFac '11/04/2016!!!!
         VencimentActual = CalculaVenciment(Venciment, DataDac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)
'         ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(DataDac) & "] Where EmpresaCodi= " & EmpCodi & " and  [NumFactura] = " & NumFac
         Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataDac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
'[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4]
         Q.rdoParameters(0) = iD '[IdFactura]
         Q.rdoParameters(1) = empCodi '[EmpresaCodi]
         Q.rdoParameters(2) = EmpSerie & CliSerie '[Serie]'11/04/2016!!!!
         Q.rdoParameters(3) = NumFac '[NumFactura]
         Q.rdoParameters(4) = Di '[DataInici]
         Q.rdoParameters(5) = Df '[DataFi]
         Q.rdoParameters(6) = DataDac '[DataFactura]
         Q.rdoParameters(7) = Now '[DataEmissio]
         Q.rdoParameters(8) = VencimentActual '[DataVenciment]
         Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
         Q.rdoParameters(10) = FacturaTotal '[Total]
         Q.rdoParameters(11) = Cli '[ClientCodi]
         Q.rdoParameters(12) = clientNom '[ClientNom]
         Q.rdoParameters(13) = ClientNif '[ClientNif]
         Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
         Q.rdoParameters(15) = clientCp '[ClientCp]
         Q.rdoParameters(16) = clientTel '[Tel]
         Q.rdoParameters(17) = clientFax '[Fax]
         Q.rdoParameters(18) = clienteMail '[email]
         Q.rdoParameters(19) = ClientLliure '[ClientLliure]
         Q.rdoParameters(20) = empNom '[EmpNom]
         Q.rdoParameters(21) = empNif '[EmpNif]
         Q.rdoParameters(22) = empAdresa '[EmpAdresa]
         Q.rdoParameters(23) = empCp '[EmpCp]
         Q.rdoParameters(24) = empTel '[EmpTel]
         Q.rdoParameters(25) = empFax '[EmpFax]
         Q.rdoParameters(26) = empEMail '[Empemail]
         Q.rdoParameters(27) = EmpLliure '[EmpLliure]
         Q.rdoParameters(28) = BaseIvaTipus_1 '[Base1]
         Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
         Q.rdoParameters(30) = BaseIvaTipus_2 '[Base2]
         Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
         Q.rdoParameters(32) = BaseIvaTipus_3 '[Base3]
         Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
         Q.rdoParameters(34) = BaseIvaTipus_4 '[Base4]
         Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
         Q.rdoParameters(36) = BaseRecTipus_1 '[Rec1]
         Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
         Q.rdoParameters(38) = BaseRecTipus_2 '[Rec2]
         Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
         Q.rdoParameters(40) = BaseRecTipus_3 '[Rec3]
         Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
         Q.rdoParameters(42) = BaseRecTipus_4 '[Rec4]
         Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
         Q.rdoParameters(44) = clientCiutat '[Rec4]
         Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
         Q.rdoParameters(46) = empCiutat '[Rec4]
         Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
         Q.rdoParameters(48) = valorIva1 '[valorIva1]
         Q.rdoParameters(49) = valorIva2 '[valorIva2]
         Q.rdoParameters(50) = valorIva3 '[valorIva3]
         Q.rdoParameters(51) = valorIva4 '[valorIva4]
         Q.rdoParameters(52) = valorRec1 '[valorRec1]
         Q.rdoParameters(53) = valorRec2 '[valorRec2]
         Q.rdoParameters(54) = valorRec3 '[valorRec3]
         Q.rdoParameters(55) = valorRec4 '[valorRec4]
         Q.rdoParameters(56) = IvaRec1 '[IvaRec1]
         Q.rdoParameters(57) = IvaRec2 '[IvaRec2]
         Q.rdoParameters(58) = IvaRec3 '[IvaRec3]
         Q.rdoParameters(59) = IvaRec4 '[IvaRec4]
         Q.rdoParameters(60) = "V1.20040304" '[Reservat]
         Q.Execute
         creaRebuts DataDac, iD, empCodi, Cli
      End If
      
      DoEvents
      
      If Abs(BaseFactura2) > 0.001 Then
          Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Group By TipusIva")
          While Not Rs.EOF
             If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                Select Case Rs("TipusIva")
                    Case 1: Tot_BaseIvaTipus_1 = Rs("Ba")
                    Case 2: Tot_BaseIvaTipus_2 = Rs("Ba")
                    Case 3: Tot_BaseIvaTipus_3 = Rs("Ba")
                    Case Else: Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_4 + Rs("Ba")
                End Select
             End If
            Rs.MoveNext
            DoEvents
        Wend
        Rs.Close
        If TipusFacturacio = 2 Then
           Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Where Acabat = 0 Group By TipusIva")
           While Not Rs.EOF
              If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                 Select Case Rs("TipusIva")
                    Case 1: Tot_BaseRecTipus_1 = Rs("Ba")
                    Case 2: Tot_BaseRecTipus_2 = Rs("Ba")
                    Case 3: Tot_BaseRecTipus_3 = Rs("Ba")
                    Case Else: Tot_BaseRecTipus_4 = Tot_BaseRecTipus_4 + Rs("Ba")
                 End Select
              End If
              Rs.MoveNext
              DoEvents
           Wend
           Rs.Close
           Tot_BaseIvaTipus_1 = Tot_BaseIvaTipus_1 - Tot_BaseRecTipus_1
           Tot_BaseIvaTipus_2 = Tot_BaseIvaTipus_2 - Tot_BaseRecTipus_2
           Tot_BaseIvaTipus_3 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_3
           Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_4
        End If
         
         FacturaContador Year(DataDac), empCodi, Cli, NumFac, NumFacNoAria '11/04/2016
         VencimentActual = CalculaVenciment(Venciment, DataDac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)

         If Not NumFacNoAria = -1 Then ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(DataDac) & "] Where EmpresaCodi= " & empCodi & " and  [NumFactura] = " & NumFacNoAria
         Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataDac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
      
        If Not (UCase(EmpresaActual) = UCase("Tena") Or Year(Di) < 2010) Then
             Tot_BaseIvaTipus_1 = 0
             BaseIvaTipus_1_Iva = 0
             Tot_BaseIvaTipus_2 = 0
             BaseIvaTipus_2_Iva = 0
             Tot_BaseIvaTipus_3 = 0
             BaseIvaTipus_3_Iva = 0
             Tot_BaseIvaTipus_4 = 0
             BaseIvaTipus_4_Iva = 0
             Tot_BaseRecTipus_1 = 0
             BaseRecTipus_1_Rec = 0
             Tot_BaseRecTipus_2 = 0
             BaseRecTipus_2_Rec = 0
             Tot_BaseRecTipus_3 = 0
             BaseRecTipus_3_Rec = 0
             Tot_BaseRecTipus_4 = 0
             BaseRecTipus_4_Rec = 0
             IvaRec1 = 0
             IvaRec2 = 0
             IvaRec3 = 0
             IvaRec4 = 0
         End If
      
      
         Q.rdoParameters(0) = "Previsio_" + iD '[IdFactura]
         Q.rdoParameters(1) = empCodi '[EmpresaCodi]
         Q.rdoParameters(2) = EmpSerie & CliSerie '[Serie]'11/04/2016!!!!
         Q.rdoParameters(3) = NumFacNoAria '[NumFactura]
         Q.rdoParameters(4) = Di '[DataInici]
         Q.rdoParameters(5) = Df '[DataFi]
         Q.rdoParameters(6) = DataDac '[DataFactura]
         Q.rdoParameters(7) = Now '[DataEmissio]
         Q.rdoParameters(8) = VencimentActual '[DataVenciment]
         Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
         Q.rdoParameters(10) = BaseFactura2 '[Total]
         Q.rdoParameters(11) = Cli '[ClientCodi]
         Q.rdoParameters(12) = clientNom '[ClientNom]
         Q.rdoParameters(13) = "" 'ClientNif '[ClientNif]
         Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
         Q.rdoParameters(15) = clientCp '[ClientCp]
         Q.rdoParameters(16) = clientTel '[Tel]
         Q.rdoParameters(17) = clientFax '[Fax]
         Q.rdoParameters(18) = clienteMail '[email]
         Q.rdoParameters(19) = ClientLliure '[ClientLliure]
         Q.rdoParameters(20) = empNom '[EmpNom]
         Q.rdoParameters(21) = "" 'EmpNif '[EmpNif]
         Q.rdoParameters(22) = empAdresa '[EmpAdresa]
         Q.rdoParameters(23) = empCp '[EmpCp]
         Q.rdoParameters(24) = empTel '[EmpTel]
         Q.rdoParameters(25) = empFax '[EmpFax]
         Q.rdoParameters(26) = empEMail '[Empemail]
         Q.rdoParameters(27) = EmpLliure '[EmpLliure]
         Q.rdoParameters(28) = Tot_BaseIvaTipus_1 'BaseIvaTipus_1 '[Base1]
         Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
         Q.rdoParameters(30) = Tot_BaseIvaTipus_2 'BaseIvaTipus_2 '[Base2]
         Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
         Q.rdoParameters(32) = Tot_BaseIvaTipus_3 'BaseIvaTipus_3 '[Base3]
         Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
         Q.rdoParameters(34) = Tot_BaseIvaTipus_4 'BaseIvaTipus_4 '[Base4]
         Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
         Q.rdoParameters(36) = Tot_BaseRecTipus_1 '[Rec1]
         Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
         Q.rdoParameters(38) = Tot_BaseRecTipus_2 '[Rec2]
         Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
         Q.rdoParameters(40) = Tot_BaseRecTipus_3 '[Rec3]
         Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
         Q.rdoParameters(42) = Tot_BaseRecTipus_4 '[Rec4]
         Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
         Q.rdoParameters(44) = clientCiutat '[Rec4]
         Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
         Q.rdoParameters(46) = empCiutat      '[Rec4]
         Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
         Q.rdoParameters(48) = valorIva1 '[valorIva1]
         Q.rdoParameters(49) = valorIva2 '[valorIva2]
         Q.rdoParameters(50) = valorIva3 '[valorIva3]
         Q.rdoParameters(51) = valorIva4 '[valorIva4]
         Q.rdoParameters(52) = valorRec1 '[valorRec1]
         Q.rdoParameters(53) = valorRec2     '[valorRec2]
         Q.rdoParameters(54) = valorRec3     '[valorRec3]
         Q.rdoParameters(55) = valorRec4     '[valorRec4]
         Q.rdoParameters(56) = IvaRec1       '[IvaRec1]
         Q.rdoParameters(57) = IvaRec2       '[IvaRec2]
         Q.rdoParameters(58) = IvaRec3       '[IvaRec3]
         Q.rdoParameters(59) = IvaRec4       '[IvaRec4]
         Q.rdoParameters(60) = "V2.20090102" '[Reservat]
   
         Q.Execute
         DoEvents
      
      End If
      
      ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataDac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio "
      
      DoEvents
      
      If DescontePp > 0 Then
         DoEvents
         ExecutaComandaSql "Update TmpFactuacio_2 Set IdFactura = 'Previsio_' + IdFactura  "
         DoEvents
         ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,3) " ' Where Import is null "
         DoEvents
         ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataDac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio_2 "
         DoEvents
      End If
      FacturaContador Year(DataDac), empCodi, Cli '11/04/2016 !!!!!!!!
   Else    ' Si no import desmarquem albarans ----------------------------------------
        D = Di
        While D <= Df
            ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set MotiuModificacio = '' Where Client = " & Cli & " And MotiuModificacio like '" & iD & "%' "
            D = DateAdd("d", 1, D)
            DoEvents
        Wend
   End If
   
    'Sincronitzacio MURANO
    Dim campNif As String
    Dim rsMurano As rdoResultset
    Dim rsPC As rdoResultset
    
    Set rsPC = Db.OpenResultset("select isnull(Valor, '') Pc from constantsempresa where camp='ProgramaContable'")
    If Not rsPC.EOF Then
        If rsPC("Pc") = "SAGE" Then
    
            campNif = "CampNif"
            If empCodi <> 0 Then campNif = empCodi & "_CampNif"
            Set rsMurano = Db.OpenResultset("select empresa, cifDni from silema_ts.sage.dbo.empresas where cifDni in (select valor collate Modern_Spanish_CI_AS from constantsempresa where camp = '" & campNif & "')")
            If Not rsMurano.EOF Then
                InsertFeineaAFer "SincroMURANOFactura", "[" & iD & "]", "[" & DataDac & "]", "[" & NumFac & "]", "[" & NomTaulaFacturaIva(DataDac) & "]"
            End If
        End If
    End If
   
   DoEvents

End Sub


Sub FacturaClientSemanal(Cli As Double, Di As Date, Df As Date, DataFac As Date, Venciment As String, Refacturar As String, CondicioArticle As String, empresa)
    Dim BaseFactura2 As Double, BaseFactura As Double, BaseIvaTipus_1 As Double, BaseIvaTipus_2 As Double, BaseIvaTipus_3 As Double, BaseIvaTipus_4 As Double, BaseIvaTipus_1_Iva As Double, BaseIvaTipus_2_Iva As Double, BaseIvaTipus_3_Iva As Double, BaseIvaTipus_4_Iva As Double, iD As String, D As Date, i As Integer
    Dim Rs As rdoResultset, Q As rdoQuery, FacturaTotal As Double, BaseRecTipus_1 As Double, BaseRecTipus_2 As Double, BaseRecTipus_3 As Double, BaseRecTipus_4  As Double, BaseRecTipus_1_Rec As Double, BaseRecTipus_2_Rec As Double, BaseRecTipus_3_Rec As Double, BaseRecTipus_4_Rec As Double, NumFacNoAria As Double
    Dim NumFac As Double, DescontePp As Double, TipusFacturacio As Double, empCodi As Double, EmpSerie As String
    Dim CliSerie As String, Tot_BaseIvaTipus_1 As Double, Tot_BaseIvaTipus_2 As Double, Tot_BaseIvaTipus_3 As Double, Tot_BaseIvaTipus_4 As Double, Tot_BaseRecTipus_1 As Double, Tot_BaseRecTipus_2 As Double, Tot_BaseRecTipus_3  As Double, Tot_BaseRecTipus_4  As Double, VencimentActual As Date, AceptaDevolucions
    Dim clientNom As String, ClientNif As String, clientAdresa As String, clientCp As String, ClientLliure As String, empNom As String, empNif As String, empAdresa As String, empCp As String, EmpLliure As String, empTel, empFax, empEMail, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCampMercantil, empCiutat, ClientNomComercial As String, Tarifa As Integer, Impostos As Double, ClientCodiFact
    Dim sql As String, valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double, IvaRec1 As Double, IvaRec2 As Double, IvaRec3 As Double, IvaRec4 As Double, PreusActuals As Boolean, DiesVenciment, DiaPagament, FormaPagoLlista, Clis() As String
    Dim BaseIBEE As Double
      
    FacturacioCreaTaulesBuides DataFac
    PreusActuals = False
    If InStr(Refacturar, "Preus Actuals") > 0 Then PreusActuals = True
    
    CarregaDadesEmpresa Cli, empCodi, EmpSerie, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa
    CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
    CarregaDadesClientVenciment Cli, DiesVenciment, DiaPagament, FormaPagoLlista
    CarregaDadesClientAgregats Cli, Clis
    If clientNom = "" Then clientNom = ClientNomComercial
    If ClientNomComercial = "" Then ClientNomComercial = clientNom
    
    iD = Format(Now, "dd-mm-yy hh:mm:ss ")
    Set Rs = Db.OpenResultset("Select newid()")
    If Not Rs.EOF Then iD = Rs(0)
    Rs.Close
    
    ExecutaComandaSql "Update Articles set TipoIva = 2 where not (TipoIva = 1 or TipoIva = 2 or TipoIva = 3 or TipoIva = 4)"
    ExecutaComandaSql "Update Articles set desconte = 1 where not desconte in (1,2,3,4)"
    
    For i = 0 To UBound(Clis)
        CarregaDadesClient Clis(i), clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
        If clientNom = "" Then clientNom = ClientNomComercial
        If ClientNomComercial = "" Then ClientNomComercial = clientNom
        FacturaClientRecullDades Val(Clis(i)), Di, Df, Refacturar, PreusActuals, iD, ClientNomComercial, CondicioArticle, DataFac
        If UBound(Clis) = 0 Then clientNom = ""
    Next
    
    CarregaDadesClient Cli, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
    If Not AceptaDevolucions Then
        ExecutaComandaSql " Update TmpFactuacio Set Tornat = 0 "
        ExecutaComandaSql " Update TmpFactuacio_2 Set Tornat = 0 "
    End If
    
    If clientNom = "" Then clientNom = ClientNomComercial
    If ClientNomComercial = "" Then ClientNomComercial = clientNom
    Tarifa = ClientTarifaEspecial(Cli)
    DescontePp = ClientDescontePp(Cli)
   
    'Repartir Blanco/Negro
    Dim agrupaAlbarans As Boolean
    agrupaAlbarans = False
    If DescontePp > 0 Then
        Set Rs = Db.OpenResultset("select ISNULL(valor, '') valor from ConstantsEmpresa where camp = 'AgrupaAlbaransDPP'")
        If Not Rs.EOF Then
            If Rs("valor") = "on" Then agrupaAlbarans = True
        End If
        Rs.Close
         
        If agrupaAlbarans Then
            sql = "Insert Into TmpFactuacio_2 "
            sql = sql & "select idfactura, max(data), '" & Cli & "', Producte, iva, rec, producteNom, acabat, preu, "
            sql = sql & "import, desconte, tipusiva, '', sum(Servit),sum(Tornat) "
            sql = sql & "From TmpFactuacio "
            sql = sql & "group by idfactura, Producte, iva, rec, producteNom, acabat, preu, import, desconte, tipusiva "
            ExecutaComandaSql sql
            ExecutaComandaSql "delete from TmpFactuacio "
            ExecutaComandaSql "insert into TmpFactuacio select * from TmpFactuacio_2"
            
            Dim compraExterna As Double
            Dim totalfactura As Double
            Dim nuevoPp As Double
                    
            compraExterna = 0
            Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f left join articlesPropietats ap on f.Producte=ap.codiArticle and ap.variable='CompraExterna' where f.client='" & Cli & "' and ap.valor='on'")
            If Not Rs.EOF Then compraExterna = Rs("import")
            If compraExterna > 0 Then
                totalfactura = 0
                Set Rs = Db.OpenResultset("select isnull(sum((preu - (preu * (desconte/100)))*(Servit-Tornat)), 0) import from TmpFactuacio f where f.client='" & Cli & "'")
                If Not Rs.EOF Then totalfactura = Rs("import")
                nuevoPp = Round((((totalfactura * (1 - DescontePp / 100)) - compraExterna) / (totalfactura - compraExterna)) * 100, 2)
                If nuevoPp < 0 Then nuevoPp = 100
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,0),Servit =  round(Servit * (" & nuevoPp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 1 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (" & nuevoPp & ")/100,2),Servit =  round(Servit * (" & nuevoPp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi left join articlesPropietats on Articles.Codi=articlesPropietats.codiArticle and articlesPropietats.variable='CompraExterna' Where Articles.EsSumable = 0 and isnull(articlesPropietats.valor,'') <> 'on' and TmpFactuacio.client = " & Cli
            Else
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,0),Servit =  round(Servit * (100-" & DescontePp & ")/100,0) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 1 and TmpFactuacio.client = " & Cli
                ExecutaComandaSql "Update TmpFactuacio Set Tornat =  round(Tornat * (100-" & DescontePp & ")/100,2),Servit =  round(Servit * (100-" & DescontePp & ")/100,2) From  TmpFactuacio Join Articles on  TmpFactuacio.Producte = Articles.Codi Where Articles.EsSumable = 0 and TmpFactuacio.client = " & Cli
            End If
            
            ExecutaComandaSql "Update t2 set t2.servit = t2.servit - t1.servit , t2.tornat = t2.tornat - t1.tornat from TmpFactuacio T1 Join TmpFactuacio_2 T2 on t2.idFactura = T1.idFactura and t1.data = t2.data and t1.client=t2.client and t1.producte = t2.producte and t1.iva = t2.iva and t1.rec = t2.rec And t1.preu = t2.preu and t1.desconte = t2.desconte and t1.referencia = t2.referencia where t2.client = " & Cli
            ExecutaComandaSql "delete TmpFactuacio_2 Where servit = 0  and  tornat = 0 "
        End If
    End If
    '~Repartir Blanco/Negro
    
    TipusFacturacio = ClientTipusFacturacio(Cli)
   
    If PreuAutomatic(Cli) Then   ' Si Preu Automatic, -> a negra preu 1
        ExecutaComandaSql "Update TmpFactuacio_2 Set Preu = Articles.Preu From TmpFactuacio_2 Join Articles on  TmpFactuacio_2.Producte = Articles.Codi "
        ' Fixem El Preu De La Tarifa Espècial  SEMPRE EL 1
        sql = "Update TmpFactuacio_2 set Preu = tarifesespecialsclients.PREU "
        sql = sql & "from TmpFactuacio_2 join tarifesespecialsclients on TmpFactuacio_2.producte = tarifesespecialsclients.codi "
        sql = sql & "and tarifesespecialsclients.client = " & Cli & " "
        ExecutaComandaSql sql
        If Tarifa > 0 Then ExecutaComandaSql "Update TmpFactuacio_2 Set Preu = TarifesEspecials.Preu From  TmpFactuacio_2 Join TarifesEspecials On  TmpFactuacio_2.Producte = TarifesEspecials.Codi And TarifesEspecials.TarifaCodi = " & Tarifa & " Join Articles On TarifesEspecials.Codi = Articles.Codi Where TmpFactuacio_2.client = " & Cli
    End If
   
    'Buscar productos con IBEE
    'sql = "update TmpFactuacio_2 set referencia='[IBEE:' + isnull(ap.valor, '0') + ']' + referencia "
    'sql = sql & "from TmpFactuacio_2 f "
    'sql = sql & "left join ArticlesPropietats ap on f.Producte=ap.codiarticle and ap.variable='ENSUCRAT' "
    'sql = sql & "where isnull(ap.valor, '') <> ''"
    'ExecutaComandaSql sql
    
    'sql = "update TmpFactuacio set referencia='[IBEE:' + isnull(ap.valor, '0') + ']' +referencia "
    'sql = sql & "from TmpFactuacio f "
    'sql = sql & "left join ArticlesPropietats ap on f.Producte=ap.codiarticle and ap.variable='ENSUCRAT' "
    'sql = sql & "where isnull(ap.valor, '') <> ''"
    'ExecutaComandaSql sql
    
    ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
    ExecutaComandaSql " Update TmpFactuacio Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,4) " ' Where Import is null "
   
    DoEvents
    
    BaseFactura = 0: BaseFactura2 = 0: BaseIvaTipus_1 = 0: BaseIvaTipus_2 = 0: BaseIvaTipus_3 = 0: BaseIvaTipus_4 = 0: FacturaTotal = 0: BaseRecTipus_1 = 0: BaseRecTipus_2 = 0: BaseRecTipus_3 = 0: BaseRecTipus_4 = 0: IvaRec1 = 0: IvaRec2 = 0: IvaRec3 = 0: IvaRec4 = 0
    'BaseIBEE = 0
    Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura = Rs(0)
    Rs.Close
    
    Set Rs = Db.OpenResultset("Select Round(Sum(Import),3) From TmpFactuacio_2 ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BaseFactura2 = Rs(0)
    Rs.Close
    
    If Abs(BaseFactura) > 0.0001 Or Abs(BaseFactura2) > 0.0001 Then
    Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Group By TipusIva")
    While Not Rs.EOF
        If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
            Select Case Rs("TipusIva")
                Case 1: BaseIvaTipus_1 = Rs("Ba")
                Case 2: BaseIvaTipus_2 = Rs("Ba")
                Case 3: BaseIvaTipus_3 = Rs("Ba")
                Case Else: BaseIvaTipus_4 = BaseIvaTipus_4 + Rs("Ba")
            End Select
        End If
        Rs.MoveNext
        DoEvents
    Wend
    Rs.Close
    
    'Set Rs = Db.OpenResultset("select isnull(cast(replace(substring(referencia,charindex('IBEE:',referencia)+5,charindex(']',referencia,charindex('IBEE:',referencia)+5)-charindex('IBEE:',referencia)-5), ',','.') as float),0) * (Servit-Tornat) IBEE from TmpFactuacio where referencia like '%IBEE%'")
    'While Not Rs.EOF
    '    BaseIBEE = BaseIBEE + Rs("IBEE")
    '    Rs.MoveNext
    '    DoEvents
    'Wend
    'Rs.Close
    
    If TipusFacturacio = 2 Then
        Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio Where Acabat = 0 Group By TipusIva")
        While Not Rs.EOF
            If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                Select Case Rs("TipusIva")
                    Case 1: BaseRecTipus_1 = Rs("Ba")
                    Case 2: BaseRecTipus_2 = Rs("Ba")
                    Case 3: BaseRecTipus_3 = Rs("Ba")
                    Case Else: BaseRecTipus_4 = BaseRecTipus_4 + Rs("Ba")
                End Select
            End If
            Rs.MoveNext
            DoEvents
        Wend
        Rs.Close
        BaseIvaTipus_1 = BaseIvaTipus_1 - BaseRecTipus_1
        BaseIvaTipus_2 = BaseIvaTipus_2 - BaseRecTipus_2
        BaseIvaTipus_3 = BaseIvaTipus_3 - BaseRecTipus_3
        BaseIvaTipus_4 = BaseIvaTipus_4 - BaseRecTipus_4
    End If
    
    CalculaIvas BaseIvaTipus_1, BaseIvaTipus_1_Iva, BaseIvaTipus_2, BaseIvaTipus_2_Iva, BaseIvaTipus_3, BaseIvaTipus_3_Iva, BaseIvaTipus_4, BaseIvaTipus_4_Iva, BaseRecTipus_1, BaseRecTipus_1_Rec, BaseRecTipus_2, BaseRecTipus_2_Rec, BaseRecTipus_3, BaseRecTipus_3_Rec, BaseRecTipus_4, BaseRecTipus_4_Rec, IvaRec1, IvaRec2, IvaRec3, IvaRec4, DataFac, False
    Impostos = BaseIvaTipus_1_Iva + BaseRecTipus_1_Rec + IvaRec1 + BaseIvaTipus_2_Iva + BaseRecTipus_2_Rec + IvaRec2 + BaseIvaTipus_3_Iva + BaseRecTipus_3_Rec + IvaRec3 + BaseIvaTipus_4_Iva + BaseRecTipus_4_Rec + IvaRec4
    
    If Not (UCase(EmpresaActual) = UCase("Tena") Or Year(Di) < 2010) Then
    
    Else
        If BaseFactura2 > 0.001 Then BaseFactura2 = BaseFactura2 + Impostos
    End If
    
    FacturaTotal = BaseIvaTipus_1 + BaseIvaTipus_2 + BaseIvaTipus_3 + BaseIvaTipus_4 _
                  + BaseRecTipus_1 + BaseRecTipus_2 + BaseRecTipus_3 + BaseRecTipus_4 _
                  + Impostos
                  '+ BaseIBEE + (BaseIBEE * 0.1)
    
    TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, DataFac
    
    If TipusFacturacio = 4 Then
        BaseIvaTipus_1 = 0
        BaseIvaTipus_2 = 0
        BaseIvaTipus_3 = 0
        BaseIvaTipus_4 = 0
    End If

    NumFac = -1
    If Abs(BaseFactura) > 0.001 Then
        FacturaContador Year(DataFac), empCodi, Cli, NumFac '11/04/2016!!!!
        VencimentActual = CalculaVenciment(Venciment, DataFac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)
        Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataFac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
        Q.rdoParameters(0) = iD '[IdFactura]
        Q.rdoParameters(1) = empCodi '[EmpresaCodi]
        Q.rdoParameters(2) = EmpSerie & CliSerie '[Serie]'11/04/2016!!!!
        Q.rdoParameters(3) = NumFac '[NumFactura]
        Q.rdoParameters(4) = Di '[DataInici]
        Q.rdoParameters(5) = Df '[DataFi]
        Q.rdoParameters(6) = DataFac '[DataFactura]
        Q.rdoParameters(7) = Now '[DataEmissio]
        Q.rdoParameters(8) = VencimentActual '[DataVenciment]
        Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
        Q.rdoParameters(10) = FacturaTotal '[Total]
        Q.rdoParameters(11) = Cli '[ClientCodi]
        Q.rdoParameters(12) = clientNom '[ClientNom]
        Q.rdoParameters(13) = ClientNif '[ClientNif]
        Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
        Q.rdoParameters(15) = clientCp '[ClientCp]
        Q.rdoParameters(16) = clientTel '[Tel]
        Q.rdoParameters(17) = clientFax '[Fax]
        Q.rdoParameters(18) = clienteMail '[email]
        Q.rdoParameters(19) = ClientLliure '[ClientLliure]
        Q.rdoParameters(20) = empNom '[EmpNom]
        Q.rdoParameters(21) = empNif '[EmpNif]
        Q.rdoParameters(22) = empAdresa '[EmpAdresa]
        Q.rdoParameters(23) = empCp '[EmpCp]
        Q.rdoParameters(24) = empTel '[EmpTel]
        Q.rdoParameters(25) = empFax '[EmpFax]
        Q.rdoParameters(26) = empEMail '[Empemail]
        Q.rdoParameters(27) = EmpLliure '[EmpLliure]
        Q.rdoParameters(28) = BaseIvaTipus_1 '[Base1]
        Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
        Q.rdoParameters(30) = BaseIvaTipus_2 '[Base2]
        Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
        Q.rdoParameters(32) = BaseIvaTipus_3 '[Base3]
        Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
        Q.rdoParameters(34) = BaseIvaTipus_4 '[Base4]
        Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
        Q.rdoParameters(36) = BaseRecTipus_1 '[Rec1]
        Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
        Q.rdoParameters(38) = BaseRecTipus_2 '[Rec2]
        Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
        Q.rdoParameters(40) = BaseRecTipus_3 '[Rec3]
        Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
        Q.rdoParameters(42) = BaseRecTipus_4 '[Rec4]
        Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
        Q.rdoParameters(44) = clientCiutat '[Rec4]
        Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
        Q.rdoParameters(46) = empCiutat '[Rec4]
        Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
        Q.rdoParameters(48) = valorIva1 '[valorIva1]
        Q.rdoParameters(49) = valorIva2 '[valorIva2]
        Q.rdoParameters(50) = valorIva3 '[valorIva3]
        Q.rdoParameters(51) = valorIva4 '[valorIva4]
        Q.rdoParameters(52) = valorRec1 '[valorRec1]
        Q.rdoParameters(53) = valorRec2 '[valorRec2]
        Q.rdoParameters(54) = valorRec3 '[valorRec3]
        Q.rdoParameters(55) = valorRec4 '[valorRec4]
        Q.rdoParameters(56) = IvaRec1 '[IvaRec1]
        Q.rdoParameters(57) = IvaRec2 '[IvaRec2]
        Q.rdoParameters(58) = IvaRec3 '[IvaRec3]
        Q.rdoParameters(59) = IvaRec4 '[IvaRec4]
        Q.rdoParameters(60) = "V1.20040304" '[Reservat]
        Q.Execute
        creaRebuts DataFac, iD, empCodi, Cli
    End If
    
    DoEvents
    
    If Abs(BaseFactura2) > 0.001 Then
        Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Group By TipusIva")
        While Not Rs.EOF
            If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                Select Case Rs("TipusIva")
                    Case 1: Tot_BaseIvaTipus_1 = Rs("Ba")
                    Case 2: Tot_BaseIvaTipus_2 = Rs("Ba")
                    Case 3: Tot_BaseIvaTipus_3 = Rs("Ba")
                    Case Else: Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_4 + Rs("Ba")
                End Select
            End If
            Rs.MoveNext
            DoEvents
        Wend
        Rs.Close
        
        If TipusFacturacio = 2 Then
            Set Rs = Db.OpenResultset("Select TipusIva,Round(Sum(Import),3) As Ba From TmpFactuacio_2 Where Acabat = 0 Group By TipusIva")
            While Not Rs.EOF
                If Not IsNull(Rs("Ba")) And Not IsNull(Rs("TipusIva")) Then
                    Select Case Rs("TipusIva")
                        Case 1: Tot_BaseRecTipus_1 = Rs("Ba")
                        Case 2: Tot_BaseRecTipus_2 = Rs("Ba")
                        Case 3: Tot_BaseRecTipus_3 = Rs("Ba")
                        Case Else: Tot_BaseRecTipus_4 = Tot_BaseRecTipus_4 + Rs("Ba")
                    End Select
                End If
                Rs.MoveNext
                DoEvents
            Wend
            Rs.Close
            Tot_BaseIvaTipus_1 = Tot_BaseIvaTipus_1 - Tot_BaseRecTipus_1
            Tot_BaseIvaTipus_2 = Tot_BaseIvaTipus_2 - Tot_BaseRecTipus_2
            Tot_BaseIvaTipus_3 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_3
            Tot_BaseIvaTipus_4 = Tot_BaseIvaTipus_3 - Tot_BaseRecTipus_4
        End If
         
        FacturaContador Year(DataFac), empCodi, Cli, NumFac, NumFacNoAria '11/04/2016
        VencimentActual = CalculaVenciment(Venciment, DataFac, Cli, DiesVenciment, DiaPagament, FormaPagoLlista)
        
        If Not NumFacNoAria = -1 Then ExecutaComandaSql "Delete [" & NomTaulaFacturaIva(DataFac) & "] Where EmpresaCodi= " & empCodi & " and  [NumFactura] = " & NumFacNoAria
        Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(DataFac) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
        
        If Not (UCase(EmpresaActual) = UCase("Tena") Or Year(Di) < 2010) Then
            Tot_BaseIvaTipus_1 = 0
            BaseIvaTipus_1_Iva = 0
            Tot_BaseIvaTipus_2 = 0
            BaseIvaTipus_2_Iva = 0
            Tot_BaseIvaTipus_3 = 0
            BaseIvaTipus_3_Iva = 0
            Tot_BaseIvaTipus_4 = 0
            BaseIvaTipus_4_Iva = 0
            Tot_BaseRecTipus_1 = 0
            BaseRecTipus_1_Rec = 0
            Tot_BaseRecTipus_2 = 0
            BaseRecTipus_2_Rec = 0
            Tot_BaseRecTipus_3 = 0
            BaseRecTipus_3_Rec = 0
            Tot_BaseRecTipus_4 = 0
            BaseRecTipus_4_Rec = 0
            IvaRec1 = 0
            IvaRec2 = 0
            IvaRec3 = 0
            IvaRec4 = 0
         End If
        
         Q.rdoParameters(0) = "Previsio_" + iD '[IdFactura]
         Q.rdoParameters(1) = empCodi '[EmpresaCodi]
         Q.rdoParameters(2) = EmpSerie & CliSerie '[Serie]'11/04/2016!!!!
         Q.rdoParameters(3) = NumFacNoAria '[NumFactura]
         Q.rdoParameters(4) = Di '[DataInici]
         Q.rdoParameters(5) = Df '[DataFi]
         Q.rdoParameters(6) = DataFac '[DataFactura]
         Q.rdoParameters(7) = Now '[DataEmissio]
         Q.rdoParameters(8) = VencimentActual '[DataVenciment]
         Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
         Q.rdoParameters(10) = BaseFactura2 '[Total]
         Q.rdoParameters(11) = Cli '[ClientCodi]
         Q.rdoParameters(12) = clientNom '[ClientNom]
         Q.rdoParameters(13) = "" 'ClientNif '[ClientNif]
         Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
         Q.rdoParameters(15) = clientCp '[ClientCp]
         Q.rdoParameters(16) = clientTel '[Tel]
         Q.rdoParameters(17) = clientFax '[Fax]
         Q.rdoParameters(18) = clienteMail '[email]
         Q.rdoParameters(19) = ClientLliure '[ClientLliure]
         Q.rdoParameters(20) = empNom '[EmpNom]
         Q.rdoParameters(21) = "" 'EmpNif '[EmpNif]
         Q.rdoParameters(22) = empAdresa '[EmpAdresa]
         Q.rdoParameters(23) = empCp '[EmpCp]
         Q.rdoParameters(24) = empTel '[EmpTel]
         Q.rdoParameters(25) = empFax '[EmpFax]
         Q.rdoParameters(26) = empEMail '[Empemail]
         Q.rdoParameters(27) = EmpLliure '[EmpLliure]
         Q.rdoParameters(28) = Tot_BaseIvaTipus_1 'BaseIvaTipus_1 '[Base1]
         Q.rdoParameters(29) = BaseIvaTipus_1_Iva '[Iva1]
         Q.rdoParameters(30) = Tot_BaseIvaTipus_2 'BaseIvaTipus_2 '[Base2]
         Q.rdoParameters(31) = BaseIvaTipus_2_Iva '[Iva2]
         Q.rdoParameters(32) = Tot_BaseIvaTipus_3 'BaseIvaTipus_3 '[Base3]
         Q.rdoParameters(33) = BaseIvaTipus_3_Iva '[Iva3]
         Q.rdoParameters(34) = Tot_BaseIvaTipus_4 'BaseIvaTipus_4 '[Base4]
         Q.rdoParameters(35) = BaseIvaTipus_4_Iva '[Iva4]
         Q.rdoParameters(36) = Tot_BaseRecTipus_1 '[Rec1]
         Q.rdoParameters(37) = BaseRecTipus_1_Rec '[BaseRec1]
         Q.rdoParameters(38) = Tot_BaseRecTipus_2 '[Rec2]
         Q.rdoParameters(39) = BaseRecTipus_2_Rec '[BaseRec1]
         Q.rdoParameters(40) = Tot_BaseRecTipus_3 '[Rec3]
         Q.rdoParameters(41) = BaseRecTipus_3_Rec '[BaseRec1]
         Q.rdoParameters(42) = Tot_BaseRecTipus_4 '[Rec4]
         Q.rdoParameters(43) = BaseRecTipus_4_Rec '[Rec4]
         Q.rdoParameters(44) = clientCiutat '[Rec4]
         Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
         Q.rdoParameters(46) = empCiutat      '[Rec4]
         Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
         Q.rdoParameters(48) = valorIva1 '[valorIva1]
         Q.rdoParameters(49) = valorIva2 '[valorIva2]
         Q.rdoParameters(50) = valorIva3 '[valorIva3]
         Q.rdoParameters(51) = valorIva4 '[valorIva4]
         Q.rdoParameters(52) = valorRec1 '[valorRec1]
         Q.rdoParameters(53) = valorRec2     '[valorRec2]
         Q.rdoParameters(54) = valorRec3     '[valorRec3]
         Q.rdoParameters(55) = valorRec4     '[valorRec4]
         Q.rdoParameters(56) = IvaRec1       '[IvaRec1]
         Q.rdoParameters(57) = IvaRec2       '[IvaRec2]
         Q.rdoParameters(58) = IvaRec3       '[IvaRec3]
         Q.rdoParameters(59) = IvaRec4       '[IvaRec4]
         Q.rdoParameters(60) = "V2.20090102" '[Reservat]
        
         Q.Execute
         DoEvents
    End If
      
    ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataFac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio "
      
    DoEvents
      
    If DescontePp > 0 Then
        DoEvents
        ExecutaComandaSql "Update TmpFactuacio_2 Set IdFactura = 'Previsio_' + IdFactura  "
        DoEvents
        ExecutaComandaSql " Update TmpFactuacio_2 Set Import = Round((Preu * (Servit - Tornat)) * (100-Desconte)/100,3) " ' Where Import is null "
        DoEvents
        ExecutaComandaSql "Insert Into [" & NomTaulaFacturaData(DataFac) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) Select [IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat] From TmpFactuacio_2 "
        DoEvents
      End If
      FacturaContador Year(DataFac), empCodi, Cli '11/04/2016 !!!!!!!!
    Else    ' Si no import desmarquem albarans ----------------------------------------
        D = Di
        While D <= Df
            ExecutaComandaSql "Update [" & DonamNomTaulaServit(D) & "] Set MotiuModificacio = '' Where Client = " & Cli & " And MotiuModificacio like '" & iD & "%' "
            D = DateAdd("d", 1, D)
            DoEvents
        Wend
    End If
   
    'Sincronitzacio MURANO
    'If UCase(EmpresaActual) = UCase("Tena") Or UCase(EmpresaActual) = UCase("Hitrs") Or UCase(EmpresaActual) = UCase("Concordia") Or UCase(EmpresaActual) = UCase("Padecava") Then
    '    InsertFeineaAFer "SincroMURANOFactura", "[" & iD & "]", "[" & DataFac & "]", "[" & NumFac & "]", "[" & NomTaulaFacturaIva(DataFac) & "]"
    'End If
   
    DoEvents
   
    EnviaFacturaEmailSecre Cli, "", CStr(NumFac), "[" & NomTaulaFacturaIva(DataFac) & "]", iD

End Sub

Function NumFacturaAdd(emp) As Double
   Dim n As Double, Rs As rdoResultset, Prefixe As String

   NumFacturaAdd = -1

   Set Rs = Db.OpenResultset("Select Camp from constantsempresa where camp = '" & emp & "_CampSeguentFactura' ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Prefixe = Left(Rs(0), InStr(Rs(0), "_"))
   Rs.Close

   Set Rs = Db.OpenResultset("Select * from ConstantsEmpresa ")
   While Not Rs.EOF
      If Not IsNull(Rs(0)) And Not IsNull(Rs(1)) Then
         If Rs(0) = Prefixe & "CampSeguentFactura" Then
            If IsNumeric(Rs(1)) Then NumFacturaAdd = Rs(1)
         End If
      End If
      Rs.MoveNext
   Wend
   Rs.Close

   If NumFacturaAdd = -1 Then NumFacturaAdd = 1

   ExecutaComandaSql "Delete ConstantsEmpresa Where Camp = '" & Prefixe & "CampSeguentFactura" & "' "
   ExecutaComandaSql "Insert Into ConstantsEmpresa (Camp,Valor) Values ('" & Prefixe & "CampSeguentFactura" & "'," & NumFacturaAdd + 1 & ") "

End Function


Function NumFacturaNegraAdd(emp) As Double
   Dim n As Double, Rs As rdoResultset, Prefixe As String
   
   NumFacturaNegraAdd = 0
   
   Set Rs = Db.OpenResultset("Select Camp from constantsempresa where camp = '" & emp & "_CampSeguentFacturaNoAria' ")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Prefixe = Left(Rs(0), InStr(Rs(0), "_"))
   Rs.Close
   
   Set Rs = Db.OpenResultset("Select * from ConstantsEmpresa ")
   While Not Rs.EOF
      If Not IsNull(Rs(0)) And Not IsNull(Rs(1)) Then
         If Rs(0) = Prefixe & "CampSeguentFacturaNoAria" Then
            If IsNumeric(Rs(1)) Then NumFacturaNegraAdd = Rs(1)
         End If
      End If
      Rs.MoveNext
   Wend
   Rs.Close
   
   If NumFacturaNegraAdd = 0 Then NumFacturaNegraAdd = -1
   ExecutaComandaSql "Delete ConstantsEmpresa Where Camp = '" & Prefixe & "CampSeguentFacturaNoAria" & "' "
   ExecutaComandaSql "Insert Into ConstantsEmpresa (Camp,Valor) Values ('" & Prefixe & "CampSeguentFacturaNoAria" & "'," & NumFacturaNegraAdd - 1 & ") "
   
End Function



Function sqlData(D As Date) As String
   
   sqlData = " CONVERT(datetime,'" & Format(D, "dd/mm/yy") & "',3) "
   
End Function


Function SqlDataMinute(D As Date) As String
   
   SqlDataMinute = " CONVERT(datetime,'" & Format(D, "dd/mm/yy hh:mm:ss") & "',3) "
   
End Function



