Attribute VB_Name = "SincroDbExterns"

Sub SincroDbExternaBdpUnDia(D As Date, PathDb As String)
    Dim DbSp As New rdoConnection, t As rdoTable, Off, f, Re, NumTic, codiProd, NomProd, Unitats, import, Preu, Hora As Date, Q As rdoQuery, Q2 As rdoQuery, Q3 As rdoQuery, botiga, D1 As Date, K, Kk, TipusVenta
Exit Sub
PathDb = "\\5.22.20.117\bdpbo\TpvHos\Datos01\TERM001"
On Error GoTo nor
    'If Not TePing("5.22.20.117") Then Exit Sub
    
    D1 = DateSerial(Year(D), Month(D), Day(D))
    If Not ExisteixTaula(NomTaulaVentas(D)) Then CreaTaulesDadesTpv D
    
    botiga = 66
    File = "LT001" & Format(D, "yymmdd") & "-" & botiga & ".DAT"
    LastData = UltimaDataGet(File)
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set Fss = FS.GetFile(PathDb & "\" & File)
    Informa2 "Vigilem " & File & " "
' \\pfarmals.portalfarma.com\descargas\Precios.txt

    If LastData < Fss.DateLastModified Then
        InformaMiss "Interpretant " & File
        UltimaDataSet File, Fss.DateLastModified
        MyMkDir AppPath
        MyMkDir AppPath & "\Tmp"
        MyKill AppPath & "\Tmp\" & File
        FileCopy PathDb & "\" & File, AppPath & "\Tmp\" & File
    
        Set Q = Db.CreateQuery("", "select * from [" & NomTaulaVentas(D) & "] where botiga = " & botiga & " and day(data) =" & Day(D) & "  and num_tick = ? and plu = ? and quantitat = ?  and import = ? ")
        Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaVentas(D) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar) Values (?,?,?,?,?,?,?,?,?,?,?) ")
        Set Q3 = Db.CreateQuery("", "Insert Into Articles (Codi,CodiGenetic,Nom,Preu,PreuMajor,Desconte,EsSumable,Familia,TipoIva,NoDescontesEspecials) Values (?,?,?,?,?,?,?,?,?,?) ")
        Kk = 0
        Set Rs = Db.OpenResultset("Select * from recordsfiles where Path = '" & File & "' ")
        If Rs.EOF Then
            ExecutaComandaSql "Delete [" & NomTaulaVentas(D) & "] where botiga = " & botiga & " and day(data) =" & Day(D) & " "
        Else
            If Not IsNull(Rs("Nom")) Then Kk = Int(Rs("nom"))
        End If
        
        f = FreeFile
        Open AppPath & "\Tmp\" & File For Input As #f
        K = 0
        While Not EOF(f)
           Re = Input(288, #f)
           K = K + 1
           
            If K > Kk Then
                NumTic = Mid(Re, 1, 9)
                codiProd = CDbl(Mid(Re, 15, 13))
                NomProd = Trim(Mid(Re, 28, 30))
                Unitats = CDbl(Mid(Re, 58, 5))
                If InStr(Re, "INVITA.") > 0 Then
                    Preu = 0
                    import = 0
                    TipusVenta = "Desc_100"
                Else
                    TipusVenta = "V"
                    Preu = CDbl(Mid(Re, 64, 9))
                    import = CDbl(Mid(Re, 74, 9))
                End If
                
                Hora = D1 + CDate(Mid(Re, 273, 5))
           'Debug.Print "NumTic = " & NumTic & " NomProd  = " & NomProd & " Unitats  = " & Unitats & " Import  = " & Import
           
                If ArticleCodiNom(codiProd) = codiProd Then
                    ExecutaComandaSql "Delete Articles Where codi = " & codiProd
                    Q3.rdoParameters(0) = codiProd
                    Q3.rdoParameters(1) = codiProd
                    Q3.rdoParameters(2) = NomProd
                    Q3.rdoParameters(3) = Preu
                    Q3.rdoParameters(4) = Preu
                    Q3.rdoParameters(5) = 1
                    Q3.rdoParameters(6) = 1
                    Q3.rdoParameters(7) = "Auto"
                    Q3.rdoParameters(8) = 0
                    Q3.rdoParameters(9) = 0
                    Q3.Execute
                    Missatges_CalEnviar "Articles", ""
                End If
           
           
'                Q.rdoParameters(0) = NumTic
'                Q.rdoParameters(1) = codiProd
'                Q.rdoParameters(2) = Unitats
'                Q.rdoParameters(3) = Import
           
'                Set rs = Q.OpenResultset()
'                If rs.EOF Then           'Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar
                    Q2.rdoParameters(0) = botiga
                    Q2.rdoParameters(1) = Hora
                    Q2.rdoParameters(2) = 1
                    Q2.rdoParameters(3) = NumTic
                    Q2.rdoParameters(4) = ""
                    Q2.rdoParameters(5) = codiProd
                    Q2.rdoParameters(6) = Unitats
                    Q2.rdoParameters(7) = import
                    Q2.rdoParameters(8) = TipusVenta
                    Q2.rdoParameters(9) = ""
                    Q2.rdoParameters(10) = ""
                    Q2.Execute
 '               End If
           End If
           DoEvents
        Wend
        ExecutaComandaSql "Delete recordsfiles where Path = '" & File & "' "
        ExecutaComandaSql "Insert Into recordsfiles (Path,Nom) Values ('" & File & "','" & K & "') "
        
    End If

nor:
    TancaComPuguis f

End Sub

Sub SincroDbExternaBoc()
    Dim DbSp As New rdoConnection, t As rdoTable, Off, f, Re, NumTic, codiProd, NomProd, Unitats, import, Preu, Hora As Date, Q As rdoQuery, Q2 As rdoQuery, Q3 As rdoQuery, botiga, D1 As Date, K, Kk, TipusVenta, D As Date

On Error GoTo nor

    File = "\\pfarmals.portalfarma.com\descargas\Precios.txt"
    FileCopy File, "c:\Bot.txt" ' AppPath & "\Tmp\" & File
    
    LastData = UltimaDataGet(File)
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set Fss = FS.GetFile(PathDb & "\" & File)
    Informa2 "Vigilem " & File & " "
    
    If LastData < Fss.DateLastModified Then
        InformaMiss "Interpretant " & File
        UltimaDataSet File, Fss.DateLastModified
        MyMkDir AppPath
        MyMkDir AppPath & "\Tmp"
        MyKill AppPath & "\Tmp\" & File
        FileCopy PathDb & "\" & File, AppPath & "\Tmp\" & File
    
        Set Q = Db.CreateQuery("", "select * from [" & NomTaulaVentas(D) & "] where botiga = " & botiga & " and day(data) =" & Day(D) & "  and num_tick = ? and plu = ? and quantitat = ?  and import = ? ")
        Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaVentas(D) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar) Values (?,?,?,?,?,?,?,?,?,?,?) ")
        Set Q3 = Db.CreateQuery("", "Insert Into Articles (Codi,CodiGenetic,Nom,Preu,PreuMajor,Desconte,EsSumable,Familia,TipoIva,NoDescontesEspecials) Values (?,?,?,?,?,?,?,?,?,?) ")
        Kk = 0
        Set Rs = Db.OpenResultset("Select * from recordsfiles where Path = '" & File & "' ")
        If Rs.EOF Then
            ExecutaComandaSql "Delete [" & NomTaulaVentas(D) & "] where botiga = " & botiga & " and day(data) =" & Day(D) & " "
        Else
            If Not IsNull(Rs("Nom")) Then Kk = Int(Rs("nom"))
        End If
        
        f = FreeFile
        Open AppPath & "\Tmp\" & File For Input As #f
        K = 0
        While Not EOF(f)
           Re = Input(288, #f)
           K = K + 1
           
            If K > Kk Then
                NumTic = Mid(Re, 1, 9)
                codiProd = CDbl(Mid(Re, 15, 13))
                NomProd = Trim(Mid(Re, 28, 30))
                Unitats = CDbl(Mid(Re, 58, 5))
                If InStr(Re, "INVITA.") > 0 Then
                    Preu = 0
                    import = 0
                    TipusVenta = "Desc_100"
                Else
                    TipusVenta = "V"
                    Preu = CDbl(Mid(Re, 64, 9))
                    import = CDbl(Mid(Re, 74, 9))
                End If
                
                Hora = D1 + CDate(Mid(Re, 273, 5))
           'Debug.Print "NumTic = " & NumTic & " NomProd  = " & NomProd & " Unitats  = " & Unitats & " Import  = " & Import
           
                If ArticleCodiNom(codiProd) = codiProd Then
                    ExecutaComandaSql "Delete Articles Where codi = " & codiProd
                    Q3.rdoParameters(0) = codiProd
                    Q3.rdoParameters(1) = codiProd
                    Q3.rdoParameters(2) = NomProd
                    Q3.rdoParameters(3) = Preu
                    Q3.rdoParameters(4) = Preu
                    Q3.rdoParameters(5) = 1
                    Q3.rdoParameters(6) = 1
                    Q3.rdoParameters(7) = "Auto"
                    Q3.rdoParameters(8) = 0
                    Q3.rdoParameters(9) = 0
                    Q3.Execute
                    Missatges_CalEnviar "Articles", ""
                End If
           
           
'                Q.rdoParameters(0) = NumTic
'                Q.rdoParameters(1) = codiProd
'                Q.rdoParameters(2) = Unitats
'                Q.rdoParameters(3) = Import
           
'                Set rs = Q.OpenResultset()
'                If rs.EOF Then           'Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar
                    Q2.rdoParameters(0) = botiga
                    Q2.rdoParameters(1) = Hora
                    Q2.rdoParameters(2) = 1
                    Q2.rdoParameters(3) = NumTic
                    Q2.rdoParameters(4) = ""
                    Q2.rdoParameters(5) = codiProd
                    Q2.rdoParameters(6) = Unitats
                    Q2.rdoParameters(7) = import
                    Q2.rdoParameters(8) = TipusVenta
                    Q2.rdoParameters(9) = ""
                    Q2.rdoParameters(10) = ""
                    Q2.Execute
 '               End If
           End If
           DoEvents
        Wend
        ExecutaComandaSql "Delete recordsfiles where Path = '" & File & "' "
        ExecutaComandaSql "Insert Into recordsfiles (Path,Nom) Values ('" & File & "','" & K & "') "
        
    End If

nor:
    TancaComPuguis f

End Sub


Sub SincronitzaFornsEnrich()
    Dim Rs As rdoResultset, rsClis As rdoResultset, rsTraspas As rdoResultset
    Dim sql As String
    Dim fecha As Date
    
On Error GoTo errSincro
        
    'CLIENTS --------------------------------------------------------------------------------------------------------------------------------------
    'Actualitza
    sql = "Update Fac_FornsEnrich.dbo.clients "
    sql = sql & "SET Nom = c.NOMBRECLIENTE, Nif = c.CIF, Adresa = c.DIRECCION1, Ciutat = c.POBLACION, Cp = c.CODPOSTAL, [Nom Llarg] = c.NOMBRECOMERCIAL "
    sql = sql & "from Fac_FornsEnrich.dbo.Clients C_Hit "
    sql = sql & "left join [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES c  on C_Hit.Codi = c.CODCLIENTE + 1000 "
    'sql = sql & "where isnull(c.cif, '') <> '' or isnull(c.NOMBRECOMERCIAL, '') <> '' "
    sql = sql & "where isnull(c.CODCONTABLE, '') <> '430000000000' and c.DESCATALOGADO = 'F' and c.NOMBRECLIENTE is not null "
    If Not ExecutaComandaSql(sql) Then Exit Sub
    
    'Inserta nous
    sql = "insert into fac_fornsEnrich.dbo.clients "
    sql = sql & "select CODCLIENTE + 1000 Codi, NOMBRECLIENTE Nom, CIF Nif, DIRECCION1 Adresa, POBLACION Ciutat, "
    sql = sql & "CODPOSTAL Cp, '' Lliure, NOMBRECOMERCIAL [Nom Llarg], 3 [Tipus Iva], 2 [Preu Base], 0 [Desconte ProntoPago], "
    sql = sql & "0 [Desconte 1], 0 [Desconte 2], 0 [Desconte 3], 0 [Desconte 4], 0 [Desconte 5], 0 AlbaraValorat "
    sql = sql & "From [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES "
    'sql = sql & "where (isnull(cif, '') <> '' or isnull(NOMBRECOMERCIAL, '') <> '') and DESCATALOGADO = 'F' and CODCLIENTE + 1000 not in (select codi from fac_fornsEnrich.dbo.clients)"
    sql = sql & "where (isnull(CODCONTABLE, '') <> '430000000000') and DESCATALOGADO = 'F' and NOMBRECLIENTE is not null and CODCLIENTE + 1000 not in (select codi from fac_fornsEnrich.dbo.clients)"
    ExecutaComandaSql sql

    'Esborra vells
    sql = "INSERT INTO Fac_FornsEnrich.dbo.Clients_Zombis "
    sql = sql & "select GETDATE(), Codi, NOM, Nif, Adresa, Ciutat, Cp, Lliure, [Nom Llarg], [Tipus Iva], [Preu Base], [Desconte ProntoPago], "
    sql = sql & "[Desconte 1] , [Desconte 2], [Desconte 3], [Desconte 4], [Desconte 5], AlbaraValorat "
    sql = sql & "From Fac_FornsEnrich.dbo.clients "
    'sql = sql & "where Codi not in (select c.CODCLIENTE + 1000 from [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES c where (isnull(c.cif, '') <> '' or isnull(c.NOMBRECOMERCIAL, '') <> '') and DESCATALOGADO = 'F' ) "
    sql = sql & "where Codi not in (select c.CODCLIENTE + 1000 from [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES c where (isnull(c.CODCONTABLE, '') <> '430000000000') and c.DESCATALOGADO = 'F' ) "
    ExecutaComandaSql sql
    
    'sql = "DELETE FROM Fac_FornsEnrich.dbo.Clients where Codi not in (select c.CODCLIENTE + 1000 from [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES c where (isnull(c.cif, '') <> '' or isnull(c.NOMBRECOMERCIAL, '') <> '') and DESCATALOGADO = 'F')"
    sql = "DELETE FROM Fac_FornsEnrich.dbo.Clients where Codi not in (select c.CODCLIENTE + 1000 from [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES c where (isnull(c.CODCONTABLE, '') <> '430000000000') and DESCATALOGADO = 'F')"
    ExecutaComandaSql sql

    'FAMILIES -------------------------------------------------------------------------------------------------------------------------------------
    'Nivell 1
    'ExecutaComandaSql "insert into families select distinct NUMDPTO Nom, 'Article' Pare, 4 Estatus, 1 Nivell, '' Utilitza from [FORNS-ENRICH].[ENRICH_MNG].[dbo].FAMILIAS"

    'Nivell 2
    'ExecutaComandaSql "insert into families select distinct NUMSECCION Nom, NUMDPTO Pare, 4 Estatus, 2 Nivell, '' Utilitza from [FORNS-ENRICH].[ENRICH_MNG].[dbo].FAMILIAS"
    
    'Nivell 3
    'ExecutaComandaSql "insert into families select distinct CAST(NUMFAMILIA AS NVARCHAR)+ '.' + DESCRIPCION Nom, NUMSECCION Pare, 4 Estatus, 3 Nivell, '' Utilitza from [FORNS-ENRICH].[ENRICH_MNG].[dbo].FAMILIAS"

    'ARTICLES --------------------------------------------------------------------------------------------------------------------------------------
    'Actualitza
    sql = "Update Fac_FornsEnrich.dbo.Articles "
    sql = sql & "SET nom = a.DESCRIPCION, "
    sql = sql & "Familia = cast(F.NUMFAMILIA as nvarchar) + '.' + F.DESCRIPCION, CodiGenetic = a.CODARTICULO, "
    sql = sql & "TipoIva = a.tipoimpuesto, NoDescontesEspecials = 1 "
    sql = sql & "from Fac_FornsEnrich.dbo.Articles A_hit "
    sql = sql & "left join [FORNS-ENRICH].[ENRICH_MNG].[dbo].ARTICULOS a  on A_hit.Codi=a.CODARTICULO "
    sql = sql & "left join [FORNS-ENRICH].[ENRICH_MNG].[dbo].FAMILIAS f on a.FAMILIA = f.NUMFAMILIA "
    sql = sql & "where not ISNULL(a.DESCRIPCION,'') = '' and a.DESCATALOGADO='F'"
    ExecutaComandaSql sql

    'Inserta nous
    sql = "INSERT INTO Fac_FornsEnrich.dbo.Articles "
    sql = sql & "select distinct a.CODARTICULO Codi, a.DESCRIPCION Nom, 1  Preu, 1 as preumajor, 1 as desconte, 1 as essumable, cast(F.NUMFAMILIA as nvarchar) + '.' + F.DESCRIPCION as familia, a.CODARTICULO CodiGenetic, a.tipoimpuesto TipoIva, 1 NodescontesEspecial "
    sql = sql & "from [FORNS-ENRICH].[ENRICH_MNG].[dbo].ARTICULOS a "
    'sql = sql & "Join [FORNS-ENRICH].[ENRICH_MNG].[dbo].PRECIOSVENTA p on p.codarticulo = a.CODARTICULO "
    sql = sql & "left join [FORNS-ENRICH].[ENRICH_MNG].[dbo].FAMILIAS f on a.FAMILIA = f.NUMFAMILIA "
    sql = sql & "where not ISNULL(a.DESCRIPCION,'') = '' and a.DESCATALOGADO='F' and a.CODARTICULO not in (select codi from Fac_FornsEnrich.dbo.Articles) "
    ExecutaComandaSql sql
    
    'Esborra vells
    sql = "INSERT INTO Fac_FornsEnrich.dbo.Articles_Zombis "
    sql = sql & "select GETDATE(), Codi, NOM, PREU, PreuMajor, Desconte, EsSumable, Familia, CodiGenetic, TipoIva, NoDescontesEspecials from Fac_FornsEnrich.dbo.Articles where Codi not in (select a.CODARTICULO from [FORNS-ENRICH].[ENRICH_MNG].[dbo].ARTICULOS a where not ISNULL(DESCRIPCION,'') = '' and a.DESCATALOGADO='F') "
    ExecutaComandaSql sql

    sql = "DELETE FROM Fac_FornsEnrich.dbo.Articles WHERE Codi not in (select a.CODARTICULO from [FORNS-ENRICH].[ENRICH_MNG].[dbo].ARTICULOS a where not ISNULL(DESCRIPCION,'') = '' and a.DESCATALOGADO='F')"
    ExecutaComandaSql sql
    
errSincro:
    
End Sub

Sub SincronitzaComandaFornsEnrich()
    Dim Rs As rdoResultset, rsClis As rdoResultset, rs2015 As rdoResultset
    Dim sql As String
    Dim fecha As Date
    
On Error GoTo errSincro
        
    'COMANDA ------------------------------------------------------------------------------------------------------------------------------------------
    fecha = Now()
    fecha = DateAdd("D", 1, fecha)
    'fecha = DateAdd("D", -6, fecha)
    
    DonamNomTaulaServit (fecha)
    
    'SELECT ALBVENTACAB.NUMSERIE, ALBVENTACAB.NUMALBARAN, ALBVENTACAB.FECHA, ALBVENTACAB.CODCLIENTE,CLIENTES.NOMBRECLIENTE,ALBVENTALIN.NUMLIN,
    'ALBVENTALIN.CODARTICULO,ALBVENTALIN.REFERENCIA,ALBVENTALIN.DESCRIPCION,ALBVENTALIN.UNIDADESTOTAL,ALBVENTALIN.PRECIO,
    'ALBVENTALIN.DTO,ALBVENTALIN.TOTAL FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB
    'LEFT  JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN ON ALBVENTACAB.NUMSERIE = ALBVENTALIN.NUMSERIE AND ALBVENTACAB.NUMALBARAN= ALBVENTALIN.NUMALBARAN AND ALBVENTACAB.N=ALBVENTALIN.N
    'LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES ON ALBVENTACAB.CODCLIENTE=CLIENTES.CODCLIENTE
    'WHERE FECHA='10/19/2016' AND NOMBRECLIENTE LIKE 'BOTIGA%'

    'Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE NOMBRECLIENTE LIKE 'BOTIGA%'")
    'Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE (isnull(CLIENTES.cif, '') <> '' AND  isnull(CLIENTES.NOMBRECLIENTE, '') <> '' AND isnull(CLIENTES.NOMBRECOMERCIAL, '') <> '') OR CLIENTES.NOMBRECLIENTE LIKE 'BOTIGA%'")
    Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE isnull(CODCONTABLE, '') <> '430000000000'")
    
    While Not rsClis.EOF
        Set rsCab = Db.OpenResultset("SELECT CAB.* , SER.DESCRIPCION SERIE_VE FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB CAB LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].SERIES SER on CAB.NUMSERIE=SER.SERIE WHERE day(FECHA)=" & Day(fecha) & " and MONTH(FECHA)=" & Month(fecha) & " AND YEAR(FECHA)=" & Year(fecha) & " AND CODCLIENTE = " & rsClis("CODCLIENTE"))
        While Not rsCab.EOF
            Set rsLin = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN WHERE NUMSERIE='" & rsCab("NUMSERIE") & "' and NUMALBARAN='" & rsCab("NUMALBARAN") & "' and N='" & rsCab("N") & "'")
            ExecutaComandaSql "Delete from Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] where client = '" & rsClis("CODCLIENTE") + 1000 & "' and comentari like '%IdAlbara:" & rsCab("NUMALBARAN") & "%'"
            While Not rsLin.EOF
                If rsLin("CODARTICULO") <> "-1" And rsLin("UNIDADESTOTAL") > 0 Then
                    sql = "insert into Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] (Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer) "
                    'sql = sql & "values ('" & rsClis("CODCLIENTE") + 1000 & "', '" & rsLin("CODARTICULO") & "', '" & rsLin("CODARTICULO") & "', isnull((select  Top 1 isnull(Viatge,'') Viatge from Fac_FornsEnrich.dbo.comandesmemotecnicperclient where codiarticle = '" & rsLin("CODARTICULO") & "' and client = '" & rsClis("CODCLIENTE") + 1000 & "' order by timestamp desc), 'Inicial'), isnull((select  Top 1 isnull(Viatge,'') Viatge from Fac_FornsEnrich.dbo.comandesmemotecnicperclient where codiarticle = '" & rsLin("CODARTICULO") & "' and client = '" & rsClis("CODCLIENTE") + 1000 & "' order by timestamp desc), 'Inicial'), " & rsLin("UNIDADESTOTAL") & ", 0, " & rsLin("UNIDADESTOTAL") & ", '', 91, 1, '[IdAlbara:" & rsLin("NUMALBARAN") & "]', '')"
                    sql = sql & "values ('" & rsClis("CODCLIENTE") + 1000 & "', '" & rsLin("CODARTICULO") & "', '" & rsLin("CODARTICULO") & "', '" & rsCab("SERIE_VE") & "', 'Inicial', " & rsLin("UNIDADESTOTAL") & ", 0, " & rsLin("UNIDADESTOTAL") & ", '', 91, 1, '[IdAlbara:" & rsLin("NUMALBARAN") & "]', '')"
                    ExecutaComandaSql sql
                End If
                rsLin.MoveNext
            Wend
            rsCab.MoveNext
        Wend

        rsClis.MoveNext
    Wend
    
    'COMANDA ANY 2015------------------------------------------------------------------------------------------------------------------------------------------
    fecha = Now()
    fecha = DateAdd("D", 1, fecha)
    fecha = DateAdd("YYYY", -1, fecha)
    'fecha = DateAdd("D", -6, fecha)
    
    Set rs2015 = Db.OpenResultset("select * from " & DonamNomTaulaServit(fecha))
    If rs2015.EOF Then
        
        'SELECT ALBVENTACAB.NUMSERIE, ALBVENTACAB.NUMALBARAN, ALBVENTACAB.FECHA, ALBVENTACAB.CODCLIENTE,CLIENTES.NOMBRECLIENTE,ALBVENTALIN.NUMLIN,
        'ALBVENTALIN.CODARTICULO,ALBVENTALIN.REFERENCIA,ALBVENTALIN.DESCRIPCION,ALBVENTALIN.UNIDADESTOTAL,ALBVENTALIN.PRECIO,
        'ALBVENTALIN.DTO,ALBVENTALIN.TOTAL FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB
        'LEFT  JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN ON ALBVENTACAB.NUMSERIE = ALBVENTALIN.NUMSERIE AND ALBVENTACAB.NUMALBARAN= ALBVENTALIN.NUMALBARAN AND ALBVENTACAB.N=ALBVENTALIN.N
        'LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES ON ALBVENTACAB.CODCLIENTE=CLIENTES.CODCLIENTE
        'WHERE FECHA='10/19/2016' AND NOMBRECLIENTE LIKE 'BOTIGA%'
    
        'Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE NOMBRECLIENTE LIKE 'BOTIGA%'")
        'Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE (isnull(CLIENTES.cif, '') <> '' AND  isnull(CLIENTES.NOMBRECLIENTE, '') <> '' AND isnull(CLIENTES.NOMBRECOMERCIAL, '') <> '') OR CLIENTES.NOMBRECLIENTE LIKE 'BOTIGA%'")
        Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE isnull(CODCONTABLE, '') <> '430000000000'")
        
        While Not rsClis.EOF
            Set rsCab = Db.OpenResultset("SELECT CAB.* , SER.DESCRIPCION SERIE_VE FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB CAB LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].SERIES SER on CAB.NUMSERIE=SER.SERIE WHERE day(FECHA)=" & Day(fecha) & " and MONTH(FECHA)=" & Month(fecha) & " AND YEAR(FECHA)=" & Year(fecha) & " AND CODCLIENTE = " & rsClis("CODCLIENTE"))
            While Not rsCab.EOF
                Set rsLin = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN WHERE NUMSERIE='" & rsCab("NUMSERIE") & "' and NUMALBARAN='" & rsCab("NUMALBARAN") & "' and N='" & rsCab("N") & "'")
                ExecutaComandaSql "Delete from Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] where client = '" & rsClis("CODCLIENTE") + 1000 & "' and comentari like '%IdAlbara:" & rsCab("NUMALBARAN") & "%'"
                While Not rsLin.EOF
                    If rsLin("CODARTICULO") <> "-1" And rsLin("UNIDADESTOTAL") > 0 Then
                        sql = "insert into Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] (Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer) "
                        'sql = sql & "values ('" & rsClis("CODCLIENTE") + 1000 & "', '" & rsLin("CODARTICULO") & "', '" & rsLin("CODARTICULO") & "', isnull((select  Top 1 isnull(Viatge,'') Viatge from Fac_FornsEnrich.dbo.comandesmemotecnicperclient where codiarticle = '" & rsLin("CODARTICULO") & "' and client = '" & rsClis("CODCLIENTE") + 1000 & "' order by timestamp desc), 'Inicial'), isnull((select  Top 1 isnull(Viatge,'') Viatge from Fac_FornsEnrich.dbo.comandesmemotecnicperclient where codiarticle = '" & rsLin("CODARTICULO") & "' and client = '" & rsClis("CODCLIENTE") + 1000 & "' order by timestamp desc), 'Inicial'), " & rsLin("UNIDADESTOTAL") & ", 0, " & rsLin("UNIDADESTOTAL") & ", '', 91, 1, '[IdAlbara:" & rsLin("NUMALBARAN") & "]', '')"
                        sql = sql & "values ('" & rsClis("CODCLIENTE") + 1000 & "', '" & rsLin("CODARTICULO") & "', '" & rsLin("CODARTICULO") & "', '" & rsCab("SERIE_VE") & "', 'Inicial', " & rsLin("UNIDADESTOTAL") & ", 0, " & rsLin("UNIDADESTOTAL") & ", '', 91, 1, '[IdAlbara:" & rsLin("NUMALBARAN") & "]', '')"
                        ExecutaComandaSql sql
                    End If
                    rsLin.MoveNext
                Wend
                rsCab.MoveNext
            Wend
    
            rsClis.MoveNext
        Wend
    End If
    
errSincro:
    
End Sub


Sub SincronitzaComandaXDiaFornsEnrich(dia As String)
    Dim Rs As rdoResultset, rsClis As rdoResultset
    Dim sql As String
    Dim fecha As Date
    
On Error GoTo errSincro
        
    'COMANDA ------------------------------------------------------------------------------------------------------------------------------------------
    fecha = Car(dia)
    
    DonamNomTaulaServit (fecha)
    
    'SELECT ALBVENTACAB.NUMSERIE, ALBVENTACAB.NUMALBARAN, ALBVENTACAB.FECHA, ALBVENTACAB.CODCLIENTE,CLIENTES.NOMBRECLIENTE,ALBVENTALIN.NUMLIN,
    'ALBVENTALIN.CODARTICULO,ALBVENTALIN.REFERENCIA,ALBVENTALIN.DESCRIPCION,ALBVENTALIN.UNIDADESTOTAL,ALBVENTALIN.PRECIO,
    'ALBVENTALIN.DTO,ALBVENTALIN.TOTAL FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB
    'LEFT  JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN ON ALBVENTACAB.NUMSERIE = ALBVENTALIN.NUMSERIE AND ALBVENTACAB.NUMALBARAN= ALBVENTALIN.NUMALBARAN AND ALBVENTACAB.N=ALBVENTALIN.N
    'LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES ON ALBVENTACAB.CODCLIENTE=CLIENTES.CODCLIENTE
    'WHERE FECHA='10/19/2016' AND NOMBRECLIENTE LIKE 'BOTIGA%'

    Set rsClis = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].CLIENTES WHERE isnull(CODCONTABLE, '') <> '430000000000'")
    
    While Not rsClis.EOF
        Set rsCab = Db.OpenResultset("SELECT CAB.* , SER.DESCRIPCION SERIE_VE FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTACAB CAB LEFT JOIN [FORNS-ENRICH].[ENRICH_MNG].[dbo].SERIES SER on CAB.NUMSERIE=SER.SERIE WHERE day(FECHA)=" & Day(fecha) & " and MONTH(FECHA)=" & Month(fecha) & " AND YEAR(FECHA)=" & Year(fecha) & " AND CODCLIENTE = " & rsClis("CODCLIENTE"))
        While Not rsCab.EOF
            Set rsLin = Db.OpenResultset("SELECT * FROM [FORNS-ENRICH].[ENRICH_MNG].[dbo].ALBVENTALIN WHERE NUMSERIE='" & rsCab("NUMSERIE") & "' and NUMALBARAN='" & rsCab("NUMALBARAN") & "' and N='" & rsCab("N") & "'")
            ExecutaComandaSql "Delete from Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] where client = '" & rsClis("CODCLIENTE") + 1000 & "' and comentari like '%IdAlbara:" & rsCab("NUMALBARAN") & "%'"
            While Not rsLin.EOF
                If rsLin("CODARTICULO") <> "-1" And rsLin("UNIDADESTOTAL") > 0 Then
                    sql = "insert into Fac_FornsEnrich.dbo.[servit-" & Format(fecha, "yy-mm-dd") & "] (Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer) "
                    sql = sql & "values ('" & rsClis("CODCLIENTE") + 1000 & "', '" & rsLin("CODARTICULO") & "', '" & rsLin("CODARTICULO") & "', '" & rsCab("SERIE_VE") & "', 'Inicial', " & rsLin("UNIDADESTOTAL") & ", 0, " & rsLin("UNIDADESTOTAL") & ", '', 91, 1, '[IdAlbara:" & rsLin("NUMALBARAN") & "]', '')"
                    ExecutaComandaSql sql
                End If
                rsLin.MoveNext
            Wend
            rsCab.MoveNext
        Wend

        rsClis.MoveNext
    Wend
    
errSincro:
    
End Sub







'
Function XX_SincroDbVendesIdentAmetller(idTasca) As Boolean
Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
Dim parametros As String, tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset
Dim tablaTmp2 As String, tablaTmp3 As String, idClientFinal
Dim idCliente, NombreCliente, idDep, NombreDep, NumTick, botiga, data
Dim connMysql As ADODB.Connection

'ACTUALIZA VENDAS IDENTIFICADAS HASTA 09-2012
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
sql = sql & "select otros from [V_Venut_2012-02] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-03] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-04] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-05] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-06] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-07] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-08] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-09] where Otros like '%cli%' group by otros ) t "
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
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-03] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-04] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-05] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-06] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-07] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-08] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-09] where Otros like '%" & idClientFinal & "%' ) t group by t.Num_tick,t.Botiga,t.data "
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




Function XX_SincroDbVendesAmetllerCheck() As Boolean
   Dim dia As Date, cos As String, Cos2 As String, Rs As rdoResultset, Rs5 As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset
   Dim sql, dd, mm, yyyy, dd2, mm2, yyyy2, nBot As Integer, nBal As Integer, vMysql, importLocal, importMysql, html
'NO SE UTILIZA ACTUALMENTE! 19/03/2013
'Repasa vendes dia anterior dels dos servidors, sino coincideixen torna a generar vendes.

On Error GoTo nor
'--------------------------------------------------------------
dia = DateAdd("d", -1, Now)
dd = Day(dia)
dd2 = Day(DateAdd("d", 1, dia))
mm = Month(dia)
mm2 = Month(DateAdd("d", 1, dia))
If Len(mm) < 2 Then mm = "0" & mm
If Len(mm2) < 2 Then mm2 = "0" & mm2
yyyy = Year(dia)
yyyy2 = Year(DateAdd("d", 1, dia))
Set Rs = Db.OpenResultset("select c.Codi,c.nom,w.Codi Wcodi  from ParamsHw w join clients c on w.Valor1 = c.Codi Order by c.nom ")
While Not Rs.EOF
    nBot = Left(Rs("codi"), 2)
    nBal = Right(Rs("codi"), 1)
    If Rs("Codi") = 518 Then
        nBal = 1
        nBot = 106
    End If
    'Ventas
    Set Rs2 = Db.OpenResultset("select isnull(Sum(import),0) V  from [" & NomTaulaVentas(dia) & "] where day(data) = " & Day(dia) & " and Botiga = " & Rs("Codi"))
    'Albarans
    Set Rs3 = Db.OpenResultset("select isnull(Sum(import),0) A  from [" & NomTaulaAlbarans(dia) & "] where day(data) = " & Day(dia) & " and Botiga = " & Rs("Codi"))
    'Mysql , sense mermes!
    sql = "select * from openquery (AMETLLER,'select sum(importeTotal) vMysql from dat_ticket_cabecera where idEmpresa=1 "
    sql = sql & "and idTienda=" & nBot & " and idBalanzaMaestra=" & nBal & " and idBalanzaEsclava=-1 and Usuario=''Comunicaciones'' and Operacion=''A'' "
    sql = sql & "and nombreBalanzaMaestra like ''-Balan%'' and timestamp>''" & yyyy & "-" & mm & "-" & dd & " 00:00:00.0000000'' "
    sql = sql & "and timestamp<''" & yyyy2 & "-" & mm2 & "-" & dd2 & " 00:00:00.0000000'' and idVendedor<>17  ') "
    Set Rs4 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then importLocal = Round(Rs2("V"), 2)
    If Not Rs3.EOF Then importLocal = importLocal + Round(Rs3("A"), 2)
    If Not Rs4.EOF Then
        If IsNull(Rs4("vMysql")) Or Rs4("vMysql") = "" Then
            importMysql = 0
        Else
            importMysql = Round(Rs4("vMysql"), 2)
        End If
    End If
    If importLocal <> importMysql Then
        InsertFeineaAFer "SincroDbVendesAmetller", "[" & Rs("codi") & "]", "[" & dd & "-" & mm & "-" & yyyy & "]", "", ""
    End If
    Rs.MoveNext
Wend
Rs.Close

Exit Function

nor:
    
    html = "<p><h3>Vendes Ametller Check</h3></p>"
    html = html & "<p><b>Botiga: </b>" & nBot & nBal & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dia & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! VendesCheck  de " & nBot & nBal & " ha fallat", html, "", ""
    
    Informa "Error : " & err.Description
    
End Function




Function SincroDbArqueigsAmetller(p1, P2, P3, P4, idTasca) As Boolean
Dim botiga As String, dia As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, tabla2 As String, sql As String
Dim Rs As rdoResultset, rsDades As rdoResultset, html As String, obre, tanca, responsable, tiqueInicial
Dim tiquetFinal, canviInicial, canviFinal, desquadre, calaix, clients, targetes, idArq As String, tmstmp
Dim connMysql As ADODB.Connection

On Error GoTo nor

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
'connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=mysql.casaametller.net;Port=3307;Database=casaametller_complements;User=hituser; Password=aM3fP6x8;Option=3;"
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=casaametller_complements;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
botiga = p1 'Cadena de tiendas separadas por ,
If botiga <> "" Then
    botiga = Replace(botiga, "[", "")
    botiga = Replace(botiga, "]", "")
End If
desde = P2
If desde <> "" Then
    desde = Replace(desde, "[", "")
    desde = Replace(desde, "]", "")
End If
dia = Day(desde)
If dia <> "" Then If Len(dia) = 1 Then dia = "0" & dia
mes = Month(desde)
If mes <> "" Then If Len(mes) = 1 Then mes = "0" & mes
anyo = Year(desde)
desde = CDate(desde)
'----------------------------------------------------------------------------------------------------
tabla = "[Fac_LaForneria].dbo.[V_Moviments_" & anyo & "-" & mes & "]"
tabla2 = "[Fac_LaForneria].dbo.[V_Venut_" & anyo & "-" & mes & "]"

sql = "SELECT DISTINCT Tipus_moviment AS expre, Data FROM  " & tabla
sql = sql & "WHERE botiga =" & botiga & " AND (day(data) = " & dia & ") AND (Tipus_moviment = 'W' OR Tipus_moviment = 'Wi') "
sql = sql & "ORDER BY Data, expre"
Set Rs = Db.OpenResultset(sql)
Do While Not Rs.EOF
    If Rs("expre") = "Wi" Then obre = Rs("data")
    Rs.MoveNext
    If Rs("expre") = "W" Then tanca = Rs("data")
    If obre <> "" And tanca <> "" Then
        sql = "select (select valor from DependentesExtes where nom='CODI_DEP' and id in (select dependenta from " & tabla & " where Data='" & tanca & "')) as responsable,"
        sql = sql & "(SELECT top 1 num_tick FROM " & tabla2 & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND botiga=" & botiga & " order by data) as tiquetinicial,"
        sql = sql & "(SELECT top 1 num_tick FROM " & tabla2 & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND botiga=" & botiga & " order by data desc) as tiquetfinal,"
        sql = sql & "(SELECT Sum(Import) as suma FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND (tipus_moviment = 'Wi') AND botiga=" & botiga & " ) as canviinicial,"
        sql = sql & "(SELECT Sum(Import) as suma FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND (tipus_moviment = 'W') AND botiga=" & botiga & " ) as canvifinal,"
        sql = sql & "(SELECT Import FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND tipus_moviment = 'J' AND botiga=" & botiga & " ) as desquadre,"
        sql = sql & "(SELECT Import FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND tipus_moviment = 'Z' AND botiga=" & botiga & " ) as calaix,"
        sql = sql & "(SELECT Import FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND tipus_moviment = 'G' AND botiga=" & botiga & " ) as clients,"
        sql = sql & "(SELECT Sum(Import) as suma FROM " & tabla & " WHERE data>='" & obre & "' and data<='" & tanca & "' AND tipus_moviment = 'O' AND botiga=" & botiga & " AND Motiu like '%targeta%') as targetes "
        Set rsDades = Db.OpenResultset(sql)
        If Not rsDades.EOF Then
           responsable = rsDades("responsable")
           tiquetInicial = rsDades("tiquetinicial")
           tiquetFinal = rsDades("tiquetfinal")
           canviInicial = rsDades("canviinicial")
           canviFinal = rsDades("canvifinal")
           desquadre = rsDades("desquadre")
           calaix = rsDades("calaix")
           clients = rsDades("clients")
           targetes = rsDades("targetes")
        End If
        'deutes
        If ExisteixTaula("CertificatDeutesAnticips") Then
            sql = "select sum(import) deute from CertificatDeutesAnticips where day(params3)='" & desde & "' and month(params3)='" & desde
            sql = sql & "' and year(params3)='" & desde & "' and Params2= " & botiga & " and params3>='" & obre & "' and params3<='" & tanca & "' "
            sql = sql & "group by params2"
            Set rsDades = Db.OpenResultset(sql)
            If Not rsDades.EOF Then deutes = rsDades("deute")
        End If
        Set rsDades = Db.OpenResultset("select newId() idArq,getdate() tmstmp")
        If Not rsDades.EOF Then
            idArq = rsDades("idArq")
            tmstmp = rsDades("tmstmp")
        End If
        sql = "INSERT INTO ca_ab_arqueigs_hit (id,timestamp,data,botiga,responsable,datainici,datafi,"
        sql = sql & "tiquetinici,tiquetfi,calaix,desquadre,clients,canviinicial,canviinicial_detall,"
        sql = sql & "canvifinal,canvifinal_detall,ingresefectiu,targetes,pagaments,deutes,cobraments,tiquetsnuls,"
        sql = sql & "pagamentshores,mermes,comentaris) VALUES ('" & idArq & "','" & Format(tmstmp, "yyyy-mm-dd hh:nn:ss") & ".000',"
        sql = sql & "'" & Format(desde, "yyyy-mm-dd hh:nn:ss") & ".000','" & botiga & "',"
        sql = sql & "'" & responsable & "','" & Format(obre, "yyyy-mm-dd hh:nn:ss") & ".000',"
        sql = sql & "'" & Format(tanca, "yyyy-mm-dd hh:nn:ss") & ".000','" & tiquetInicial & "','" & tiquetFinal & "',"
        sql = sql & "'" & calaix & "','" & desquadre & "','" & clients & "','" & canviInicial & "','',"
        sql = sql & "'" & canviFinal & "','','" & targetes & "','','','','','','','','')"
        'Set rsDades = Db.OpenResultset(Sql)
        connMysql.Execute sql
    End If
    Rs.MoveNext
Loop
connMysql.Close
Set connMysql = Nothing
    
'----------------------------------------------------------------------------------------------------
'----  CONTROL ERROR
'----------------------------------------------------------------------------------------------------
nor:
    If err.Number <> 0 Then
        html = "<p><h3>Resum</h3></p>"
        html = html & "<p><b>Botiga: </b>" & botiga & "</p>"
        html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
        html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
        If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio d'arqueigs de " & codiBotiga & " ha fallat", html, "", ""
        html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
        sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio d'arqueigs de " & codiBotiga & " ha fallat", html, "", ""
        
        Informa "Error : " & err.Description
    End If
End Function



Function XX_SincroDbEmailAmetller(p1, P2, P3, P4, idTasca) As Boolean
Dim sql As String, Rs As rdoResultset, html As String, proc As String, procAnt As String, fecha, log As String
'NO SE UTILIZA ACTUALMENTE! 19/03/2013
'----------------------------------------------------------------------------------------------------
'----  IMPRIME TABLA LOG Y ENVIA POR EMAIL
'----------------------------------------------------------------------------------------------------

On Error GoTo noree
    
    procAnt = ""
    html = "<p><h3>Resum migracio horaria</h3></p>"
    html = html & "<p><b>Data inici sincronitzacio: </b>" & p1 & "</p>"
    sql = "select proceso,fecha,txt from sincro_procedures_log where fecha>='" & p1 & " ' order by fecha "
    Set Rs = Db.OpenResultset(sql)
    Do While Not Rs.EOF
        proc = Rs("proceso")
        fecha = Rs("fecha")
        log = DameValor(Rs, "txt")
        If procAnt <> proc Then
            If InStr(UCase(proc), "ARTICLE") > 1 Then html = html & "<p><h4>Articles</h4></p>"
            If InStr(UCase(proc), "DEPENDENTES") > 1 Then html = html & "<p><h4>Dependentes</h4></p>"
            If InStr(UCase(proc), "TARIFES") > 1 Then html = html & "<p><h4>Tarifes</h4></p>"
            If InStr(UCase(proc), "BOTIGUES") > 1 Then html = html & "<p><h4>Botigues</h4></p>"
            If InStr(UCase(proc), "CLIENTS") > 1 Then html = html & "<p><h4>Botigues</h4></p>"
            procAnt = proc
        End If
        html = html & "<p>" & log & "</p>"
        Rs.MoveNext
    Loop
    html = html & "<p><b>Data fi sincronitzacio: </b>" & fecha & "</p>"
    If P2 <> "" Then
        sf_enviarMail "secrehit@hit.cat", P2, "Resum de migracio horaria " & p1, html, "", ""
    Else
        sf_enviarMail "secrehit@hit.cat", EmailGuardia, "Resum de migracio horaria " & p1, html, "", ""
        sf_enviarMail "secrehit@hit.cat", "programador@casaametller.net", "Resum de migracio horaria " & p1, html, "", ""
    End If
noree:

End Function




Sub TancaComPuguis(f)
On Error Resume Next
    Close #f
End Sub

Sub SincroDbExternaBdpFichajes(PathDb As String)
    Dim DbSp As New rdoConnection, t As rdoTable, Off, f, Re, NumTic, codiProd, NomProd, Unitats, import, Preu, Hora As Date, Q As rdoQuery, Q2 As rdoQuery, Q3 As rdoQuery, botiga, D1 As Date, K, Kk, D As Date, codiDep
    Dim Coditreb As Double, NomTreb As String, sHi                 As String, sHf As String, Hi As Date, Hf As Date
On Error GoTo nor

    
    If Not TePing("5.22.20.117") Then Exit Sub
    
    botiga = 66
    File = "TPVT01-202.DAT"
    LastData = UltimaDataGet(File)
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set Fss = FS.GetFile(PathDb & "\..\" & File)
    Informa2 "Vigilem " & File & " "
        
    If LastData < Fss.DateLastModified Then
        InformaMiss "Interpretant " & File
        UltimaDataSet File, Fss.DateLastModified
        MyMkDir AppPath
        MyMkDir AppPath & "\Tmp"
        MyKill AppPath & "\Tmp\" & File
        FileCopy PathDb & "\..\" & File, AppPath & "\Tmp\" & File
    
        Set QIns = Db.CreateQuery("", "Insert into CdpDadesFichador (Id,TmSt,Accio,Usuari,Idr,Lloc,Comentari) values (0,?,?,?,newid(),'" & botiga & "','[Desde:TECLADO]')")
        Set Qdel = Db.CreateQuery("", "Delete      CdpDadesFichador Where Id = 0 And TmSt = ? And Accio = ? And Usuari = ? And Lloc = '" & botiga & "' And Comentari = '[Desde:TECLADO]' ")
        
        Kk = 0
        Set Rs = Db.OpenResultset("Select * from recordsfiles where Path = '" & File & "' ")
        If Rs.EOF Then
            'ExecutaComandaSql "Delete [" & NomTaulaVentas(D) & "] where botiga = " & Botiga & " and day(data) =" & Day(D) & " "
        Else
            If Not IsNull(Rs("Nom")) Then Kk = Int(Rs("nom"))
        End If
        
        f = FreeFile
        Open AppPath & "\Tmp\" & File For Input As #f
        K = 0
        While Not EOF(f)
           Re = Input(288, #f)
           K = K + 1
            If K > Kk Then
                Coditreb = Mid(Re, 14, 5)
                NomTreb = Trim(Mid(Re, 19, 40))
                sHi = Left(Mid(Re, 59, 15), 10) & " " & Right(Mid(Re, 59, 15), 5)
                sHf = Left(Mid(Re, 115, 15), 10) & " " & Right(Mid(Re, 115, 15), 5)
           
                Set Rs = Db.OpenResultset("select d.codi  As Codi  from dependentesextes e join dependentes D on d.codi = e.id and e.nom = 'CodiBdp'  and e.valor = '" & Coditreb & "' ")
                If Rs.EOF Then
                    ExecutaComandaSql "Delete dependentesextes Where nom = 'CodiBdp'  and valor = '" & Coditreb & "' "
                    Set Rs = Db.OpenResultset("select * From Dependentes where Nom = '" & NomTreb & "' ")
                    If Rs.EOF Then
                        ExecutaComandaSql "Insert into Dependentes (Codi,nom ) Select max(codi)+1 as codi , '" & NomTreb & "' as nom  from Dependentes"
                        Set Rs = Db.OpenResultset("select * From Dependentes where Nom = '" & NomTreb & "' ")
                    End If
                    codiDep = Rs("Codi")
                    ExecutaComandaSql "Delete dependentesextes Where nom = 'CodiBdp'  and valor = '" & Coditreb & "' "
                    ExecutaComandaSql "Insert Into dependentesextes (Id,Nom,Valor) Values ('" & codiDep & "','CodiBdp','" & Coditreb & "') "
                    Set Rs = Db.OpenResultset("select d.codi As Codi from dependentesextes e join dependentes D on d.codi = e.id and e.nom = 'CodiBdp'  and e.valor = '" & Coditreb & "' ")
                End If
                codiDep = Rs("Codi")
                
                If IsDate(sHi) Then
                   Hi = CVDate(sHi)
                   Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaHoraris(Hi) & "] Where Botiga = ? And Data= ? and Dependenta = ? And Operacio= ? ")
                   Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaHoraris(Hi) & "] (Botiga,Data,Dependenta,Operacio) Values (?,?,?,?)")
                   Q1.rdoParameters(0) = botiga
                   Q1.rdoParameters(1) = Hi
                   Q1.rdoParameters(2) = codiDep
                   Q1.rdoParameters(3) = "E"  '  "P"
                   Q1.Execute
                   Q2.rdoParameters(0) = botiga
                   Q2.rdoParameters(1) = Hi
                   Q2.rdoParameters(2) = codiDep
                   Q2.rdoParameters(3) = "E"  '  "P"
                   Q2.Execute
                   
                   Qdel.rdoParameters(0) = Hi
                   Qdel.rdoParameters(1) = 1
                   Qdel.rdoParameters(2) = codiDep
                   Qdel.Execute
                   
                   QIns.rdoParameters(0) = Hi
                   QIns.rdoParameters(1) = 1
                   QIns.rdoParameters(2) = codiDep
                   QIns.Execute
                   
                End If
                Hf = Hi
                If IsDate(sHf) Then
                   Hf = CVDate(sHf)
                   Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaHoraris(Hf) & "] Where Botiga = ? And Data= ? and Dependenta = ? And Operacio= ? ")
                   Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaHoraris(Hf) & "] (Botiga,Data,Dependenta,Operacio) Values (?,?,?,?)")
                   Q1.rdoParameters(0) = botiga
                   Q1.rdoParameters(1) = Hf
                   Q1.rdoParameters(2) = codiDep
                   Q1.rdoParameters(3) = "P"
                   Q1.Execute
                   Q2.rdoParameters(0) = botiga
                   Q2.rdoParameters(1) = Hf
                   Q2.rdoParameters(2) = codiDep
                   Q2.rdoParameters(3) = "P"
                   Q2.Execute
                   
                   Qdel.rdoParameters(0) = Hf
                   Qdel.rdoParameters(1) = 2
                   Qdel.rdoParameters(2) = codiDep
                   Qdel.Execute
                   
                   QIns.rdoParameters(0) = Hf
                   QIns.rdoParameters(1) = 2
                   QIns.rdoParameters(2) = codiDep
                   QIns.Execute
                End If
            End If
           DoEvents
        Wend
        Close #f
        
        ExecutaComandaSql "Delete recordsfiles where Path = '" & File & "' "
        ExecutaComandaSql "Insert Into recordsfiles (Path,Nom) Values ('" & File & "','" & K & "') "
        
    End If

nor:

End Sub


Sub SincroDbExternaIcg(Server, NomDb, User, Psw)
    Dim Dbr As New rdoConnection, LastSincro As Date, LastSincroNew As Date, Qi As rdoQuery, Qd As rdoQuery
    Dim algun
    Dim Nnom As String
    
'Exit Sub
    InformaMiss "Sincro Icg"
    If Not TePing(CStr(Server)) Then Exit Sub
'    Dim i, DbIcg As New rdoConnection
''select codarticulo,descripcion,0,0,1,1,'',codarticulo,1,0 from [ICG_ARTICULOS]
'
'    DbIcg.Connect = "WSID=Extern;UID=Secre;PWD=Secre1234;Database=DbfRest;Server='5.164.51.150';Driver={SQL Server};DSN='';"
'    PathDb = "\\5.22.20.117\bdpbo\TpvHos\Datos01\TERM001"
    
On Error GoTo nor
    InformaEstat frmSplash.lblVersion, "Conectant Server " & Server
    Dbr.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
    Dbr.EstablishConnection rdDriverNoPrompt
    
    
'DbR.Connect = "DSN=Sp;Uid=Temp;Pwd=hit;SERVER=192.168.240.3;"
'DbR.Connect = "Provider=MSDASQL;DRIVER={MySQL ODBC 3.51 Driver};SERVER=192.168.240.3;DATABASE=Visual FoxPro Database;UID=myusername;PWD=mypassword;"
'DbR.EstablishConnection
    
    InformaEstat frmSplash.lblVersion, "Conectant !!"
    
'    LastSincro = DateAdd("Y", -10, Now)
'    Set rs = Db.OpenResultset("Select * From Records Where Concepte = 'SincroDbExternaIcgArticles'")
'    If Not rs.EOF Then LastSincro = rs("TimeStamp")
'    LastSincroNew = LastSincro
'    Algun = False
''    Set Rs = DbR.OpenResultset("SELECT     CODARTICULO, DESCRIPCION, FECHAMODIFICADO From ARTICULOS Where FECHAMODIFICADO > '" & LastSincro & "'")
'
''    SincroTb Dbr, "ARTICULOS"   '<---- Taula de usuaris
'
'    Dim cc
''    Set Rs = Dbr.OpenResultset("SELECT Top 1 *  From ARTICULOS ")
''    For Each cc In Rs.rdoColumns
''        Debug.Print cc.Name
''    Next
''inicia la sincronizacion de los productos y materias primas
'
'    Set rs = Dbr.OpenResultset("SELECT     DPTO,CODARTICULO, DESCRIPCION, FECHAMODIFICADO From ARTICULOS Where datediff(s, " & SqlDataMinute(LastSincro) & ",FECHAMODIFICADO ) >1 ")
''    Set Rs = Dbr.OpenResultset("SELECT     * From ARTICULOS  ")
'
'    While Not rs.EOF
'        If rs("FECHAMODIFICADO") > LastSincroNew Then LastSincroNew = rs("FECHAMODIFICADO")
'        Nnom = ""
'        If Not IsNull(rs("DESCRIPCION")) Then Nnom = rs("DESCRIPCION")
'        Nnom = Join(Split(Nnom, "'"), "`")
'        Nnom = Join(Split(Nnom, Chr(34)), "`")
'
'        If rs("DPTO") < 5 And rs("DPTO") > 0 Then
'           Set Rs2 = Db.OpenResultset("Select Codi From  Articles Where Codi = '" & rs("CODARTICULO") & "' ")
'           If Rs2.EOF Then
'                ExecutaComandaSql "Insert Into Articles ([Codi], [NOM], [PREU], [PreuMajor], [Desconte], [EsSumable], [Familia], [CodiGenetic], [TipoIva], [NoDescontesEspecials] ) values ('" & rs("CODARTICULO") & "','" & Nnom & "',1,1,1,1,'',1,1,1)"
'           Else
'                ExecutaComandaSql "Update Articles Set [NOM] = '" & Nnom & "' Where Codi= '" & rs("CODARTICULO") & "' "
'           End If
'        Else
'           If rs("DPTO") = 5 Then
'                Dim Idm As String
'                Set Rs2 = Db.OpenResultset("Select Id From  ccmateriasprimas Where Codigo = '" & rs("CODARTICULO") & "' ")
'                If Rs2.EOF Then
'                    ExecutaComandaSql "Insert Into ccmateriasprimas (Id,Codigo,Nombre) Values (newid(),'" & rs("CODARTICULO") & "','" & Nnom & "') "
'                Else
'                    ExecutaComandaSql "Update  ccmateriasprimas Set Nombre = '" & Nnom & "' ,Codigo = '" & rs("CODARTICULO") & "'  Where Id = '" & Rs2("Id") & "' "
'                End If
'           End If
'        End If
'
'        InformaEstat frmSplash.lblVersion, Nnom
'        rs.MoveNext
'        If Not Algun And Not rs.EOF Then
'            ExecutaComandaSql "Select * Into [ArticlesBk" & Now & "] from  Articles "
'            ExecutaComandaSql "Select * Into [ccmateriasprimasBk" & Now & "] from  ccmateriasprimas "
'        End If
'        Algun = True
'    Wend
'    If Not LastSincroNew = LastSincro Then
'        ExecutaComandaSql "Delete Records Where Concepte = 'SincroDbExternaIcgArticles' "
'        ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('SincroDbExternaIcgArticles','" & LastSincroNew & "')"
'    End If
'
  ' Terminamos con productos y materias primas y pasamos a sincronizar Almacenes
  

'    Set rs = Dbr.OpenResultset("SELECT * From ALMACEN  ")
'
'    While Not rs.EOF
'        Nnom = ""
'        If Not IsNull(rs("NOMBREALMACEN")) Then Nnom = rs("NOMBREALMACEN")
'        Nnom = Join(Split(Nnom, "'"), "`")
'        Nnom = Join(Split(Nnom, Chr(34)), "`")
'
'        Set Rs2 = Db.OpenResultset("Select nombre From ccalmacenes Where nombre = '" & rs("NOMBREALMACEN") & "' ")
'           If Rs2.EOF Then
'                ExecutaComandaSql "Insert Into ccalmacenes ([id], [nombre], [descripcion], [alta], [activo] ) values ( newid() ,'" & rs("NOMBREALMACEN") & "','" & rs("CODALMACEN") & ",'" & LastSincro & "','1')"
'           Else
'                ExecutaComandaSql "Update Articles Set [nombre] = '" & Nnom & "' Where nombre= '" & rs("NOMBREALMACEN") & "' "
'           End If
'
'        InformaEstat frmSplash.lblVersion, Nnom
'        rs.MoveNext
'        If Not Algun And Not rs.EOF Then
'            ExecutaComandaSql "Select * Into [AlmacenesBk" & Now & "] from  Almacenes "
'        End If
'        Algun = True
'    Wend
'
    LastSincro = DateAdd("d", -30, Now)
    Set Rs = Db.OpenResultset("Select * From Records Where Concepte = 'SincroDbExternaIcgTickets'")
    If Not Rs.EOF Then LastSincro = Rs("TimeStamp")
    LastSincroNew = LastSincro
    algun = False
    
    D = LastSincro
    LastSincroNew = LastSincro
                    
    While D < DateAdd("d", 1, Now)

'delete [v_venut_2011-07] where DAY(data) = 15
'insert into [v_venut_2011-07]
'select 3,hora,codvendedor,numero,'',codarticulo,unidades,precioiva,'V',0,0 from [5.164.51.150].dbfrest.dbo.TIQUETSLIN where year(hora) = 2011 and  month(hora) = 7  and  day(hora) = 15
        
        Set Rs = Dbr.OpenResultset("select numero,codarticulo,unidades,precioiva,codvendedor,hora from TIQUETSLIN where year(hora) = " & Year(D) & " and  month(hora) = " & Month(D) & "  and  day(hora) = " & Day(D) & "  and hora > CONVERT(datetime,'" & Day(LastSincro) & "/" & Month(LastSincro) & "/" & Right(Year(LastSincro), 2) & " " & Hour(LastSincro) & ":" & Minute(LastSincro) & ":" & Second(LastSincro) & "',3) order by hora   ")
        If Not Rs.EOF Then
            Set Qd = Db.CreateQuery("", "Delete [" & NomTaulaVentas(CVDate(D)) & "] where Botiga = ? And data=? And dependenta = ? And Num_tick= ? And plu = ? And Quantitat= ? and Import = ? ")
            Set Qi = Db.CreateQuery("", "Insert into [" & NomTaulaVentas(CVDate(D)) & "] (Botiga,data,dependenta,Num_tick,estat,plu,Quantitat,Import,Tipus_venta,FormaMarcar,Otros) values (?,?,?,?,'',?,?,?,'V',0,0)")
        End If
        While Not Rs.EOF
            If Rs("hora") > LastSincroNew Then LastSincroNew = Rs("hora")
            Qd.rdoParameters(0) = 3
            Qd.rdoParameters(1) = Rs("hora")
            Qd.rdoParameters(2) = Rs("codvendedor")
            Qd.rdoParameters(3) = Rs("numero")
            Qd.rdoParameters(4) = Rs("codarticulo")
            Qd.rdoParameters(5) = Rs("unidades")
            Qd.rdoParameters(6) = Rs("precioiva")
            Qd.Execute
            
            Qi.rdoParameters(0) = 3
            Qi.rdoParameters(1) = Rs("hora")
            Qi.rdoParameters(2) = Rs("codvendedor")
            Qi.rdoParameters(3) = Rs("numero")
            Qi.rdoParameters(4) = Rs("codarticulo")
            Qi.rdoParameters(5) = Rs("unidades")
            Qi.rdoParameters(6) = Rs("precioiva")
            Qi.Execute
            

            Rs.MoveNext
        Wend
        D = DateAdd("d", 1, D)
        InformaMiss "Sincro Icg Data " & D
    Wend
    
    If Not LastSincroNew = LastSincro Then
        ExecutaComandaSql "Delete Records Where Concepte = 'SincroDbExternaIcgTickets' "
        ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('SincroDbExternaIcgTickets','" & LastSincroNew & "')"
    End If
    
nor:

    InformaEstat frmSplash.lblVersion, ""
    
End Sub



Function SincroDbExternaSp(PathDb, P2, P3, P4) As Boolean
    Dim DbSp As New rdoConnection, t As rdoTable
   SincroDbExternaSp = False
On Error Resume Next
    DbSp.Close
'    PathDb = "\\192.168.240.3\c$\GrupoSP\FAE10R01\DBF03"

' \\pfarmals.portalfarma.com\descargas\Precios.txt
    UltimPathDb = PathDb
    If PathDb = "" Then
        SincroDbExternaSp = True
        Exit Function
    End If
    
    DbSp.Connect = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & UltimPathDb & "\;Exclusive=No"
    
    DbSp.EstablishConnection 'rdDriverNoPrompt
    
    SincroTb DbSp, "albclic"  '<---- Taula Clients
    SincroTb DbSp, "albclil"  '<---- Taula Clients
    SincroTb DbSp, "albclit"  '<---- Taula Clients
    
    SincroTb DbSp, "Clientes"  '<---- Taula Clients
    
    SincroTb DbSp, "Agentes"   '<---- Taula de usuaris
    SincroTb DbSp, "AgentesC"  '<---- Taula de usuaris Claus Gmail i usuaris Gmail
    SincroTb DbSp, "Clientes"  '<---- Taula Clients
    
'    For Each T In DbSp.rdoTables
'        SincroTb DbSp, T.Name
'    Next
    
    SincroDbExternaSp = True
    Exit Function
nor:

End Function

Function SincroDbExternaBdp(PathDb As String, P2, P3, P4) As Boolean
    Dim i
                           
    PathDb = "\\5.22.20.117\bdpbo\TpvHos\Datos01\TERM001"

    
    SincroDbExternaBdpUnDia Now, PathDb
    SincroDbExternaBdpUnDia DateAdd("d", -1, Now), PathDb
    SincroDbExternaBdpFichajes PathDb
    
'    For i = -1 To -30 Step -1
'        SincroDbExternaBdpUnDia DateAdd("d", i, Now), PathDb
'    Next

    
End Function

Function UltimaDataGet(Clau) As Date
    Dim Rs As rdoResultset
    
    ExecutaComandaSql "CREATE TABLE Records ([TimeStamp] [datetime] Null,[Concepte] [nvarchar] (255) NULL) ON [PRIMARY]"
    
    Set Rs = Db.OpenResultset("Select * From Records Where Concepte = '" & Clau & "'")
    If Rs.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('" & Clau & "',DATEADD(Day, -1, GetDate()))"
    Set Rs = Db.OpenResultset("Select * From Records Where Concepte = '" & Clau & "'")
    UltimaDataGet = Rs("TimeStamp")


End Function


Sub UltimaDataSet(Clau, D)
        
    ExecutaComandaSql "Update Records Set [TimeStamp] = convert(datetime,'" & Format(D, "dd/mm/yy") & "',3)+convert(datetime,'" & Hour(D) & ":" & Minute(D) & ":" & Second(D) & "',8) Where Concepte = '" & Clau & "' "
    
End Sub



Sub SincroTb(Dbr As rdoConnection, Taula As String)
    Dim c As rdoColumn, sql As String, Q As rdoQuery, Qdel As rdoQuery, LlistaCamps, LlistaInts, i, K, Rs As rdoResultset, n, Tipus, Rs2 As rdoResultset, Rs3 As rdoResultset, Codi
    Dim Guser, Gpassword
    
On Error Resume Next

    ExecutaComandaSql "Drop Table [SP_" & Taula & "] "
    sql = ""
    LlistaCamps = ""
    LlistaInts = ""
    i = 0
    
    Set Rs = Dbr.OpenResultset("Select * from " & Taula)
    If Rs.EOF Then Exit Sub
    
    For Each c In Rs.rdoColumns
        If Not sql = "" Then sql = sql & ","
        If Not LlistaCamps = "" Then LlistaCamps = LlistaCamps & ","
        If Not LlistaInts = "" Then LlistaInts = LlistaInts & ","
        
        LlistaCamps = LlistaCamps & c.Name & " "
        LlistaInts = LlistaInts & " ? "
        Tipus = " [nvarchar] (255) NULL "
        Select Case c.Type
            Case 2
                Tipus = " numeric (18,3) NULL "
            Case 9
                Tipus = " datetime NULL "
            Case Else
                Tipus = " [nvarchar] (255) NULL "
        End Select
        
        sql = sql & c.Name & " " & " " & Tipus & " "
        i = i + 1
        DoEvents
    Next
    sql = "Create table [SP_" & Taula & "]  (" & sql & ") "
    ExecutaComandaSql sql
    n = 0
    
    ExecutaComandaSql "Delete [SP_" & Taula & "] "
    Set Q = Db.CreateQuery("", "Insert Into [SP_" & Taula & "] (" & LlistaCamps & ") Values (" & LlistaInts & ") ")
    While Not Rs.EOF
        For K = 0 To i - 1
            Q.rdoParameters(K) = ""
            Q.rdoParameters(K) = Trim(Rs(K))
        Next
        Q.Execute
        Rs.MoveNext
        n = n + 1
        DoEvents
    Wend
    
    Select Case Taula
        Case "Agentes"
            Set Rs = Db.OpenResultset("Select * from SP_Agentes ")
            While Not Rs.EOF
                Codi = 1
                Set rs1 = Db.OpenResultset("Select max(codi) from Dependentes ")
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then Codi = rs1(0) + 1
                Set rs1 = Db.OpenResultset("Select * from Dependentes Where Tid = '" & Rs("ccodage") & "' ")
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then Codi = rs1(0)
                ExecutaComandaSql "Delete Dependentes Where Tid = '" & Rs("ccodage") & "' "
                ExecutaComandaSql "Insert Into Dependentes (Codi,nom,memo,telefon,[adreça],Tid) Values (" & Codi & ",'" & Rs("capeage") & "','" & Rs("cnbrage") & "','" & Rs("ctfoage") & "','" & Rs("cdirage") & "','" & Rs("ccodage") & "') "
                Rs.MoveNext
            Wend
        Case "AgentesC"
            Set Rs = Db.OpenResultset("Select * from SP_AgentesC ")
            While Not Rs.EOF
                Codi = 1
                Set rs1 = Db.OpenResultset("Select max(codi) from Dependentes ")
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then Codi = rs1(0) + 1
                Set rs1 = Db.OpenResultset("Select * from Dependentes Where Tid = '" & Rs("ccodage") & "' ")
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then Codi = rs1(0)
'GmUser:natfoodiberica.mattia@gmail.com   GmPsw:19811981  GmConactos:Mios  GmChat:Si  GmCalendar:Si
                If InStr(Rs("Notas"), "GmUser:") > 0 And InStr(Rs("Notas"), "GmPsw:") > 0 Then
                    Dim P, Pp, st, St1, St2
                    P = InStr(Rs("Notas"), "GmUser:")
                    Pp = InStr(P, Rs("Notas"), Chr(13))
                    Guser = Mid(Rs("Notas"), P + 7, Pp - P - 8)
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'uGmail' "
                    ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'uGmail','" & Guser & "') "
                    
                    P = InStr(Rs("Notas"), "GmPsw:")
                    Pp = InStr(P, Rs("Notas"), Chr(13))
                    Guser = Mid(Rs("Notas"), P + 6, Pp - P - 6)
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'passGmail' "
                    ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'passGmail','" & Guser & "') "
                    
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'sincAgenda' "
                    If InStr(Rs("Notas"), "GmCalendar:") > 0 Then ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'sincAgenda','1') "
                    
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'sincContactos' "
                    If InStr(Rs("Notas"), "GmConactos:Si") > 0 Then ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'sincContactos','1') "
                    
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'sincContactosPropis' "
                    If InStr(Rs("Notas"), "GmConactos:Mios") > 0 Then ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'sincContactosPropis','" & Rs("ccodage") & "') "
                    
                    ExecutaComandaSql "Delete DependentesExtes Where id = '" & Codi & "' And Nom = 'permisChat' "
                    If InStr(Rs("Notas"), "GmChat:Si") > 0 Then ExecutaComandaSql "Insert Into DependentesExtes (Id,Nom,Valor) Values (" & Codi & ",'permisChat','1') "
                    
                End If
                Rs.MoveNext
            Wend
        Case "Clientes"
            ExecutaComandaSql "Delete Clients"
            ExecutaComandaSql "insert into clients select cast(ccodcli as numeric) as codi,isnull(cnomcli,'') as nom,isnull(cdnicif,'')  as Nif,isnull(cdircli,'')  as adresa,isnull(cpobcli,'')  as Ciutat,isnull(cptlcli,'')  as Cp,isnull(cobscli,'')  as Lliure,isnull(cnomcom,'')  as [Nom Llarg],1 as [Tipus Iva],0 as [Preu Base],0 as [Desconte ProntoPago],0 as [Desconte 1],0 as [Desconte 2],0 as [Desconte 3],0 as [Desconte 4],0 as [Desconte 5],1 as AlbaraValorat from sp_clientes"
            ExecutaComandaSql "Delete constantsclient"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Tel' as Variable , isnull(ctfo1cli,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Movil' as Variable , isnull(ctfo2cli,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Fax' as Variable , isnull(cfaxcli,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'P_Contacte' as Variable1 , isnull(ccontacto,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'CompteCorrent' as Variable ,isnull(centidad+cagencia+ccuenta,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'DiaPagament' as Variable ,isnull(ndia1pago,0) Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Morosidad' as Variable ,isnull(nmorosos,0) Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Riesgo' as Variable ,isnull(nriesgo,0) Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Riesgo2' as Variable ,isnull(nriesgoalc,0) Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'eMail' as Variable ,isnull(email,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Grup_client' as Variable ,isnull(ccodgrup,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'Bloqueado' as Variable ,isnull(lbloqueado,'') Valor from  sp_clientes"
            ExecutaComandaSql "insert into constantsclient select  cast(ccodcli as  Numeric) as codi , 'ccodage' as Variable ,isnull(ccodage,'') Valor from  sp_clientes"

        
    End Select
    ExecutaComandaSql "Insert into ImportatsSp (TmSt,Taula,Registres) values (getdate(),'" & Taula & "'," & n & ")"
    Informa2 Taula & " --> " & n & " Registres Ok."
    Debug.Print Taula & " --> " & n & " Registres Ok."
    
End Sub


