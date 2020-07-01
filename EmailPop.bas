Attribute VB_Name = "EmailPop"
Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public emailFactura_XML As String
Sub actualizaPrevisiones(empresa As String, emailDe As String, HTMLText As String)
    Dim posInput As Integer
    Dim posName As Integer, nameInput As String, nameInputSplit() As String
    Dim posValue As Integer, valueTxt As String
    Dim textHTML As String
    Dim botiga As String
    Dim data As Date
    Dim torn As String
    Dim prevision As Double
    Dim sql As String
    
    textHTML = HTMLText
    
    On Error GoTo nor
    
    posInput = InStr(textHTML, "P_")
    
    While posInput > 0
        textHTML = Mid(textHTML, posInput)
        posName = InStr(textHTML, """")
        nameInput = Left(textHTML, posName - 1)
        If InStr(nameInput, "_") Then
            nameInputSplit = Split(nameInput, "_")
            botiga = nameInputSplit(1)
            data = CDate(nameInputSplit(2))
            torn = nameInputSplit(3)
        End If
        
        posValue = InStr(textHTML, ">")
        textHTML = Mid(textHTML, posValue + 1)
        valueTxt = Left(textHTML, InStr(textHTML, "<") - 1)
        If IsNumeric(valueTxt) Then
            prevision = CDbl(valueTxt)
            ExecutaComandaSql "delete from [" & NomTaulaMovi(data) & "] where day(data)=" & Day(data) & " and botiga=" & botiga & " and tipus_moviment='" & torn & "_C'"
            ExecutaComandaSql "insert into [" & NomTaulaMovi(data) & "] (botiga, data, dependenta, tipus_moviment, import, motiu) values (" & botiga & ", convert(datetime, '" & data & "', 103), 0, '" & torn & "_C', " & prevision & ", '" & emailDe & "')"
        End If
        
        posInput = InStr(textHTML, "P_")
    Wend
    Exit Sub
    
nor:
    posInput = posInput
End Sub

Function AsignaFacturaNumSerie(Text As String, empresa As String) As String
    Dim nFactura As String, NPRODUCTE As String, nSerie As String, P As Integer, i As Integer
    Dim D As Date, rs As rdoResultset, iD As String
            
    nFactura = 0
    NPRODUCTE = 0
    nSerie = 0
    
    
    AsignaFacturaNumSerie = "Asignem Numero de serie...<Br>"
    AsignaFacturaNumSerie = AsignaFacturaNumSerie & "Empresa: " & empresa & "<Br>"
    Text = UCase(Text)
    P = InStr(Text, "NFACTURA")
    If P > 0 Then
        nFactura = Trim(Mid(Text, P + 10, InStr(P + 10, Text, Chr(13)) - (P + 10)))
    End If
    AsignaFacturaNumSerie = AsignaFacturaNumSerie & "NFactura: " & nFactura & "<Br>"
    
    P = InStr(Text, "NPRODUCTE")
    If P > 0 Then
        NPRODUCTE = Trim(Mid(Text, P + 11, InStr(P + 11, Text, Chr(13)) - (P + 11)))
    End If
    AsignaFacturaNumSerie = AsignaFacturaNumSerie & "NPRODUCTE: " & NPRODUCTE & "<Br>"
    
    P = InStr(Text, "NSERIE")
    If P > 0 Then
        nSerie = Trim(Mid(Text, P + 8, InStr(P + 11, Text, Chr(13)) - (P + 8)))
    End If
    AsignaFacturaNumSerie = AsignaFacturaNumSerie & "NSERIE: " & nSerie & "<Br>"
    
    
    iD = ""
    For i = 1 To 12
        D = DateSerial(Year(Now), i, 1)
        If ExisteixTaula(NomTaulaFacturaIva(D)) Then
            Set rs = Db.OpenResultset("select * from " & empresa & ".dbo.[" & NomTaulaFacturaIva(D) & "] where numfactura = '" & nFactura & "'")
            If Not rs.EOF Then
                iD = rs("IdFactura")
                AsignaFacturaNumSerie = AsignaFacturaNumSerie & "Client: " & BotigaCodiNom(rs("clientcodi")) & "<Br>"
                AsignaFacturaNumSerie = AsignaFacturaNumSerie & "DataFactura: " & rs("DataFactura") & "<Br>"
                Exit For
            End If
        End If
    Next
    
    If iD = "" Then
        AsignaFacturaNumSerie = AsignaFacturaNumSerie & "Factura No Trobada en l any " & Year(Now) & " !!! <Br>"
    Else
        Set rs = Db.OpenResultset("select * from " & empresa & ".dbo.[" & NomTaulaFacturaData(D) & "] where Producte = '" & NPRODUCTE & "'")
        If Not rs.EOF Then
        
        
        End If
        
    End If
    
    
'NFACTURA: 257
'NPRODUCTE: 20213
'NSERIE: 2017110174


End Function

Sub DoEventsSleep()
    DoEvents
    Sleep 10
End Sub


Sub InterpretaFacturaPdf(File, empresa As String, emailDe As String)
    Dim URL As String
    Dim resposta As String, articlesResposta As String, errorText As String
    Dim i As Integer, P As Integer
    Dim rs As rdoResultset, rsId As rdoResultset
    Dim sql As String
    Dim aNif() As String, aBaseImponible() As String, aData() As String
    Dim NifEmisor As String, NifReceptor As String, nFactura As String, proveidorCodi As String, clientCodi As String
    Dim clientNom As String, clientAdresa As String, clientCiutat As String, clientCp As String, clientTel As String, clientFax As String, clienteMail As String
    Dim provNom As String, provAdresa As String, provCP As String, provTel As String, provFax As String, proveMail As String, provCiutat As String
    Dim baseIva21 As Double, baseIva10 As Double, baseIva4 As Double, import As Double
    Dim Iva21 As Double, Iva10 As Double, iva4 As Double, Total As Double
    Dim baseIva21_Pdf As Double, baseIva10_Pdf As Double, baseIva4_Pdf As Double
    Dim Iva21_Pdf As Double, Iva10_Pdf As Double, Iva4_Pdf As Double, total_Pdf As Double
    Dim dataFactura As Date
    Dim mpCodigo As String, mpNombre As String, mpDescripcion As String, nLote As String, cadStr As String, caducidad As String
    Dim pPrecio As String, pCantidad As String
    Dim idFactura As String, idPedido As String

    Debug.Print File
    
On Error GoTo nor

    URL = "http://www.gestiondelatienda.com/facturacion/elforn/file/getPdfFile.asp?"
    URL = URL & "file=" & File
    resposta = llegeigHtml(URL)
    
    aData = Split(Mid(Trim(Replace(Split(resposta, "FECHA")(1), ".", "")), 1, 8), "-")
    dataFactura = CDate(aData(0) & "/" & aData(1) & "/" & aData(2))

    nFactura = Split(Trim(Replace(Split(resposta, "FACTURA")(1), ".", "")), " ")(0)
    
    aNif = Split(Trim(Split(resposta, "N.I.F.")(1)), " ")
    NifEmisor = Replace(Trim(aNif(0)), ".", "")
    i = 1
    While Len(NifEmisor) < 9
        NifEmisor = NifEmisor & Trim(aNif(i))
        i = i + 1
    Wend

    proveidorCodi = ""
    Set rs = Db.OpenResultset("Select * from " & empresa & ".dbo.ccProveedores where replace(nif, '-', '') = replace('" & NifEmisor & "', '-', '')")
    If Not rs.EOF Then
        proveidorCodi = rs("id")
        
        provNom = rs("nombre")
        provAdresa = rs("direccion")
        provCP = rs("cp")
        provTel = rs("tlf1")
        provFax = rs("fax")
        proveMail = rs("eMail")
        provCiutat = rs("ciudad")
    End If

    aNif = Split(Trim(Split(resposta, "N.I.F.")(2)), " ")
    NifReceptor = Replace(Trim(aNif(0)), ".", "")
    i = 1
    While Len(NifReceptor) < 9
        NifReceptor = NifReceptor & Trim(aNif(i))
        i = i + 1
    Wend

    clientCodi = "0"
    Set rs = Db.OpenResultset("Select * from " & empresa & ".dbo.ConstantsEmpresa where replace(Valor, '-', '')=replace('" & NifReceptor & "', '-', '')")
    If Not rs.EOF Then
        If InStr(rs("camp"), "_") Then
            clientCodi = Split(rs("Camp"), "_")(0)
        End If
    End If
    
    If clientCodi = "0" Then
        Set rs = Db.OpenResultset("select * from " & empresa & ".dbo.ConstantsEmpresa where Camp like 'Camp%' order by camp")
    Else
        Set rs = Db.OpenResultset("select * from " & empresa & ".dbo.ConstantsEmpresa where Camp like '" & clientCodi & "_Camp%' order by camp")
    End If
    While Not rs.EOF
        If InStr(rs("camp"), "CampNom") Then
            clientNom = rs("valor")
        ElseIf InStr(rs("camp"), "CampAdresa") Then
            clientAdresa = rs("valor")
        ElseIf InStr(rs("camp"), "CampCiutat") Then
            clientAdresa = clientAdresa & rs("valor")
        ElseIf InStr(rs("camp"), "CampCiutat") Then
            clientAdresa = clientAdresa & " " & rs("valor")
            clientCiutat = rs("valor")
        ElseIf InStr(rs("camp"), "CampProvincia") Then
            clientAdresa = clientAdresa & " " & rs("valor")
        ElseIf InStr(rs("camp"), "CampCodiPostal") Then
            clientCp = rs("valor")
        ElseIf InStr(rs("camp"), "CampTel") Then
            clientTel = rs("valor")
        ElseIf InStr(rs("camp"), "CampFax") Then
            clientFax = rs("valor")
        ElseIf InStr(rs("camp"), "CampMail") Then
            clienteMail = rs("valor")
        End If
        rs.MoveNext
    Wend

    
    baseIva21 = 0
    Iva21 = 0
    baseIva10 = 0
    Iva10 = 0
    baseIva4 = 0
    iva4 = 0
    Total = 0

    Set rs = Db.OpenResultset("select newid() as idFactura")
    idFactura = rs("idFactura")
    
    'Recepción
    Set rs = Db.OpenResultset("select *, i.Iva PctIva from " & empresa & ".dbo.ccMateriasPrimas m left join " & DonamTaulaTipusIva(dataFactura) & " i on  m.iva = i.Tipus where m.proveedor='" & proveidorCodi & "' and m.codigo<>'' order by m.codigo")
    While Not rs.EOF
        articlesResposta = resposta
        If InStr(articlesResposta, rs("codigo")) Then
            mpCodigo = rs("codigo")
            mpNombre = rs("nombre")
            mpDescripcion = rs("descripcion")
            articlesResposta = Trim(Split(articlesResposta, mpCodigo)(1))
            If InStr(articlesResposta, mpDescripcion) Then
                articlesResposta = Trim(Mid(articlesResposta, Len(mpDescripcion) + 1))
            Else
                'Dejan 29 carácteres, como máximo, para el nombre
                articlesResposta = Trim(Mid(articlesResposta, 30))
            End If

            pCantidad = Replace(Split(articlesResposta, " ")(1), ",", ".")
            pPrecio = Replace(Split(articlesResposta, " ")(5), ",", ".")
            nLote = ""
            If Split(articlesResposta, " ")(7) = "Lote/" Then
                nLote = Split(articlesResposta, " ")(8)
            End If
            
            Set rsId = Db.OpenResultset("select newid() as idPedido")
            idPedido = rsId("idPedido")
            
            If IsNumeric(pPrecio) And IsNumeric(pCantidad) Then
                import = Round(pPrecio * pCantidad, 2)
                If nLote = "" Then nLote = LotInit(empresa, mpCodigo)
            
                cadStr = rs("Caducidad")
                If cadStr = "" Then cadStr = "1 Mes"
                P = InStr(cadStr, " ")
                If P > 0 And cadStr <> " " And cadStr <> "No" Then
                    caducidad = Left(DateAdd(IIf(UCase(Mid(cadStr, P + 1, 1)) = "A", "yyyy", Mid(cadStr, P + 1, 1)), Left(cadStr, P), Now), 10)
                End If
            
                sql = "insert into " & empresa & ".dbo.ccPedidos (id, materiaPrima, Proveedor, almacen, cantidad, fecha, recepcion, precio, activo, confirmado) values "
                sql = sql & "('" & idPedido & "', '" & rs("id") & "', '" & proveidorCodi & "', '', " & pCantidad & ", "
                sql = sql & "'" & dataFactura & "', '" & dataFactura & "', " & pPrecio & ", 1, null)"
                ExecutaComandaSql sql

                sql = "insert into " & empresa & ".dbo.ccRecepcion (Id, proveedor, matPrima, albaran, pedido, temperatura, caract, envas, usuario, fecha, aceptado, lote, facturado, caducidad) values "
                sql = sql & "(newid(), '" & proveidorCodi & "', '" & rs("id") & "', '" & nFactura & "', '" & idPedido & "', "
                sql = sql & "0, 1, 1, 'SINCRO', '" & dataFactura & "', 0, '" & nLote & "', 0, '" & caducidad & "')"
                ExecutaComandaSql sql
                
                If rs("Iva") = 1 Then
                    sql = "insert into " & tablaFacturaProformaDataE(dataFactura, empresa) & " (IdFactura, Data, Client, Producte, ProducteNom, Acabat, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) values "
                    sql = sql & " ('" & idFactura & "', '" & dataFactura & "', null, '" & rs("id") & "', '" & mpNombre & "', null, " & pPrecio & ", " & import & ", 0, 1, 4, 0, '60000000', " & pCantidad & ", 0)"
                    ExecutaComandaSql sql
                    
                    baseIva4 = baseIva4 + import
                    iva4 = iva4 + import * (rs("PctIva") / 100)
                End If
    
                If rs("Iva") = 2 Then
                    sql = "insert into " & tablaFacturaProformaDataE(dataFactura, empresa) & " (IdFactura, Data, Client, Producte, ProducteNom, Acabat, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) values "
                    sql = sql & " ('" & idFactura & "', '" & dataFactura & "', null, '" & rs("id") & "', '" & mpNombre & "', null, " & pPrecio & ", " & import & ", 0, 2, 10, 0, '60000000', " & pCantidad & ", 0)"
                    ExecutaComandaSql sql
                    
                    baseIva10 = baseIva10 + import
                    Iva10 = Iva10 + import * (rs("PctIva") / 100)
                End If
    
                If rs("Iva") = 3 Then
                    sql = "insert into " & tablaFacturaProformaDataE(dataFactura, empresa) & " (IdFactura, Data, Client, Producte, ProducteNom, Acabat, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) values "
                    sql = sql & " ('" & idFactura & "', '" & dataFactura & "', null, '" & rs("id") & "', '" & mpNombre & "', null, " & pPrecio & ", " & import & ", 0, 3, 21, 0, '60000000', " & pCantidad & ", 0)"
                    ExecutaComandaSql sql
                    
                    baseIva21 = baseIva21 + import
                    Iva21 = Iva21 + import * (rs("PctIva") / 100)
                End If
             Else
                errorText = errorText & "Càrrega errònia del producte: " & mpCodigo & "  " & mpNombre
            End If
            
        End If
        rs.MoveNext
    Wend

    Total = baseIva21 + Iva21 + baseIva10 + Iva10 + baseIva4 + iva4

    'Factura
    sql = " insert into " & tablaFacturaProformaE(dataFactura, empresa) & " values ('" & idFactura & "', '" & nFactura & "', '" & proveidorCodi & "', '', '" & dataFactura & "', '" & dataFactura & "', '" & dataFactura & "', getdate(), '" & dataFactura & "', '', " & Total & ", "
    sql = sql & " '" & clientCodi & "', '" & clientCodi & "', '" & clientNom & "', '" & NifReceptor & "', '" & clientAdresa & "', '" & clientCp & "', '" & clientTel & "', '" & clientFax & "', '" & clienteMail & "', "
    sql = sql & "'', '" & clientCiutat & "', '" & provNom & "', '" & NifEmisor & " ', '" & provAdresa & "', '" & provCP & "', '" & provTel & "', '" & provFax & "', '" & proveMail & "', '', '" & provCiutat & "', '', "
    sql = sql & baseIva4 & ", " & iva4 & ", " & baseIva10 & ", " & Iva10 & ", " & baseIva21 & ", " & Iva21 & ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "
    sql = sql & "4, 10, 21, 0, 0.5, 1.4, 5.2, 0, 0, 0, 0, 0, 'SINCRO')"
    ExecutaComandaSql sql


    i = 1
    baseIva21_Pdf = 0
    Iva21_Pdf = 0
    aBaseImponible = Split(resposta, "BASE IMPONIBLE.")
    If InStr(aBaseImponible(i), "TIPO IMPOSITIVO. 21,00 %") Then
        baseIva21_Pdf = CDbl(Replace(Split(aBaseImponible(i), "TIPO IMPOSITIVO. 21,00 %")(0), ",", "."))
        Iva21_Pdf = CDbl(Replace(Split(aBaseImponible(i), "I.V.A.")(1), ",", "."))
        i = i + 1
    End If
    
    baseIva10_Pdf = 0
    Iva10_Pdf = 0
    If InStr(aBaseImponible(i), "TIPO IMPOSITIVO. 10,00 %") Then
        baseIva10_Pdf = CDbl(Replace(Split(aBaseImponible(i), "TIPO IMPOSITIVO. 10,00 %")(0), ",", "."))
        Iva10_Pdf = CDbl(Replace(Replace(Split(aBaseImponible(i), "I.V.A.")(1), ",", "."), " ", ""))
        i = i + 1
    End If
    
    baseIva4_Pdf = 0
    Iva4_Pdf = 0
    If InStr(aBaseImponible(i), "TIPO IMPOSITIVO. 4,00 %") Then
        baseIva4_Pdf = CDbl(Replace(Split(aBaseImponible(i), "TIPO IMPOSITIVO. 4,00 %")(0), ",", "."))
        Iva4_Pdf = CDbl(Replace(Replace(Split(Trim(Split(aBaseImponible(i), "I.V.A.")(1)), " ")(0), ",", "."), " ", ""))
    End If
    
    total_Pdf = baseIva21_Pdf + Iva21_Pdf + baseIva10_Pdf + Iva10_Pdf + baseIva4_Pdf + Iva4_Pdf
        
    'Enviar email confirmación
    sf_enviarMail "Secrehit@gmail.com", "ana@solucionesit365.com", "Factura " & nFactura & " importada.", "Total Calculado: " & Total & "; Total Pdf: " & total_Pdf, "", ""
    
    Exit Sub
 
nor:
    Debug.Print err.Description
    EnviaEmail "ana@solucionesit365.com", "ERROR IMPORTANDO FACTURA: " & err.Description

End Sub


Function LotInit(empresa As String, article As String) As String
    Dim LotApodo As String
    Dim proposat As String
    Dim rs As rdoResultset
    Dim produccio As Integer
    
    ExecutaComandaSql "alter table " & empresa & ".dbo.[AppLots] ALTER COLUMN plu nvarchar(255)"

    proposat = Right("00" & Year(Now), 2) & Right("0" & DatePart("W", Now, vbMonday), 1) & Right("00" & DatePart("ww", Now, vbMonday, vbFirstFourDays), 2)
    Set rs = Db.OpenResultset("Select count (distinct Nom) Dif From " & empresa & ".dbo.[AppLots] Where Left(Nom,5) = '" & proposat & "'  ")

    produccio = 1
    If Not rs.EOF Then If Not IsNull(rs("Dif")) Then produccio = rs("Dif") + 1
    proposat = proposat & Right("000" & CStr(produccio), 3)
    
    Set rs = Db.OpenResultset("Select top 1 * from " & empresa & ".dbo.[AppLots] Where Plu = '" & article & "' order By DataI Desc ")
    If rs.EOF Then
        LotApodo = "A"
    Else
        LotApodo = rs("Apodo")
        ' La Lletra la canviem sempre
        If LotApodo = "Z" Then
            LotApodo = "A"
        Else
            LotApodo = Chr(Asc(LotApodo) + 1)
        End If
            
        If Day(rs("DataI")) = Day(Now) And Month(rs("DataI")) = Month(Now) And Year(rs("DataI")) = Year(Now) Then
            proposat = rs("Nom")
        End If
    End If


    ExecutaComandaSql "Insert Into " & empresa & ".dbo.[AppLots] ([tipus],[Nom],[Apodo],[dataI],[Plu]) Values ('Materia Prima','" & proposat & "','" & LotApodo & "',getdate(),'" & article & "') "
    Set rs = Db.OpenResultset("Select top 1 * from " & empresa & ".dbo.[AppLots] Where Plu = '" & article & "' order By DataI Desc ")
    If Not rs.EOF Then
        LotInit = rs("Nom") & LotApodo
    Else
        LotInit = ""
    End If

End Function

Sub preparaCodigoDeAccionAutorizacion(HTMLText As String, emailDe As String, empresa As String)
    Dim HTMLAux As String
    Dim codigoDeAccion As String
    Dim tagCdA As String, tagAutorizacion As String
    
    On Error GoTo nor:
    
    HTMLAux = UCase(HTMLText)
    
    tagCdA = "CODIGO_ACCION:["
    tagAutorizacion = "NAME=""FACTURA_AUTORIZADA"
    
    If InStr(HTMLAux, tagCdA) Then
        codigoDeAccion = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA), 100)
        codigoDeAccion = Mid(codigoDeAccion, 1, InStr(codigoDeAccion, "]") - 1)
        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param10 = '" & emailDe & "' where idCodigo='" & codigoDeAccion & "'"
        
        HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagAutorizacion) + Len(tagAutorizacion))
        
        If InStr(HTMLAux, ">OK<") Then
           
            ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param3 = 'OK' where idCodigo='" & codigoDeAccion & "'"
            
        End If
        
        ExecutaComandaSql "insert into feinesafer (Id, Tipus, Ciclica, Param1, tmStmp) values (newid(), 'CodigoDeAccion', 0, '" & codigoDeAccion & "', getdate())"
    End If
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR preparaCodigoDeAccionAutorizacion", "ERROR: " & err.Description, "", ""

End Sub
Sub preparaCodigoDeAccionIncidencia(HTMLText As String, emailDe As String, empresa As String)
    Dim HTMLAux As String
    Dim codigoDeAccion As String
    Dim tagCdA As String, tagComentario As String, tagIn As String, tagOut As String
    Dim rsInc As rdoResultset, rsCodigo As rdoResultset
    Dim rsOtros As rdoResultset
    Dim txtInc As String
    Dim InOut As String
    
    On Error GoTo nor:
    
    HTMLAux = UCase(HTMLText)
    
    tagCdA = "CODIGO_ACCION:["
    tagComentario = "NAME=""TD_OBSERVACIONES"
    tagIn = "NAME=""TD_IN"
    tagOut = "NAME=""TD_OUT"
    
    If InStr(HTMLAux, tagCdA) Then
        codigoDeAccion = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA), 100)
        codigoDeAccion = Mid(codigoDeAccion, 1, InStr(codigoDeAccion, "]") - 1)
        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param10 = '" & emailDe & "' where idCodigo='" & codigoDeAccion & "'"
        
        
        If InStr(HTMLAux, tagComentario) > 1 Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagComentario))
            txtInc = Mid(HTMLAux, Len(tagComentario & """>") + 1, InStr(HTMLAux, "</TD></TR>") - Len(tagComentario & """>") - 1)
        Else
            txtInc = Mid(HTMLAux, 1, 1000)
        End If
                
        Set rsCodigo = Db.OpenResultset("select * from  " & taulaCodigosDeAccion() & "  where idCodigo = '" & codigoDeAccion & "'")
        If Not rsCodigo.EOF Then
            Set rsInc = Db.OpenResultset("select * from incidencias where id=" & rsCodigo("param2"))
            If Not rsInc.EOF Then
            
                If InStr(HTMLAux, tagIn) > 1 Then
                    HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagIn))
                    InOut = Mid(HTMLAux, Len(tagIn & """>") + 1, InStr(HTMLAux, "</TD></TR>") - Len(tagIn & """>") - 1)
                    If InOut <> "" Then ExecutaComandaSql "insert into HW_CLIENTS (incidencia, codigoAccion, usuario, InOut, serie) values ('" & rsCodigo("param2") & "', '" & codigoDeAccion & "', '" & emailDe & "', 'IN', '" & InOut & "')"
                End If
                
                If InStr(HTMLAux, tagOut) > 1 Then
                    HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagOut))
                    InOut = Mid(HTMLAux, Len(tagOut & """>") + 1, InStr(HTMLAux, "</TD></TR>") - Len(tagOut & """>") - 1)
                    If InOut <> "" Then ExecutaComandaSql "insert into HW_CLIENTS (incidencia, codigoAccion, usuario, InOut, serie) values ('" & rsCodigo("param2") & "', '" & codigoDeAccion & "', '" & emailDe & "', 'OUT', '" & InOut & "')"
                End If
            
                'ExecutaComandaSql "update incidencias set estado='Resuelta' where id=" & rsCodigo("param2")
                ExecutaComandaSql "insert into Inc_Historico (id, timestamp, usuario, incidencia, tipo) values (" & rsCodigo("param2") & ", getdate(), '" & emailDe & "', '" & txtInc & "', 'TEXTO')"
                Set rsOtros = Db.OpenResultset("SELECT * FROM Inc_Link_Otros WHERE ID=" & rsCodigo("param2"))
                If Not rsOtros.EOF Then
                    ExecutaComandaSql "insert into " & rsOtros("Empresa") & ".dbo.Inc_Historico (id, timestamp, usuario, incidencia, tipo) values (" & rsOtros("idOtro") & ", getdate(), '" & emailDe & "', '" & txtInc & "', 'TEXTO')"
                End If
            Else
                ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param3 = 'KO' where idCodigo='" & codigoDeAccion & "'"
            End If
        Else
            ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param3 = 'KO' where idCodigo='" & codigoDeAccion & "'"
        End If
        
        'ExecutaComandaSql "insert into feinesafer (Id, Tipus, Ciclica, Param1, tmStmp) values (newid(), 'CodigoDeAccion', 0, '" & codigoDeAccion & "', getdate())"
    End If
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR preparaCodigoDeAccionIncidencia", "ERROR: " & err.Description, "", ""

End Sub

Sub InterpretaIncidenciaCERBERO(HTMLText As String, emailDe As String, empresa As String)
    Dim HTMLAux As String
    Dim codigoDeAccion As String
    Dim tagCdA As String
    Dim rsInc As rdoResultset, rsCodigo As rdoResultset, rsCreador As rdoResultset, rsHist As rdoResultset
    Dim incOrigen As String, incDestino As String, empOrigen As String, empDestino As String, items As String
    Dim cliDestino As String, respDestino As String
    Dim usuarioNom As String, fechaHistorico As String
    Dim sql As String, idInc As String
    Dim creador As String 'De momento será CERBERO
    
    On Error GoTo nor:
    
    HTMLAux = UCase(HTMLText)
    
    tagCdA = "CODIGO_ACCION:["
    
    If InStr(HTMLAux, tagCdA) Then
        codigoDeAccion = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA), 100)
        codigoDeAccion = Mid(codigoDeAccion, 1, InStr(codigoDeAccion, "]") - 1)
        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param10 = '" & emailDe & "' where idCodigo='" & codigoDeAccion & "'"
        
        Set rsCodigo = Db.OpenResultset("select * from  " & taulaCodigosDeAccion() & "  where idCodigo = '" & codigoDeAccion & "'")
        If Not rsCodigo.EOF Then
            empOrigen = rsCodigo("Param1")
            incOrigen = rsCodigo("Param2")
            incDestino = rsCodigo("Param3")
            empDestino = rsCodigo("Param4")
            cliDestino = rsCodigo("Param5")
            respDestino = rsCodigo("Param6")
            items = rsCodigo("Param7") 'Items modificados
            fechaHistorico = rsCodigo("Param8") 'Fecha histórico origen para poder recuperar la modificación
            
            If incDestino = "" Then 'ES UNA TAREA NUEVA Y NO HAY CORRESPONDENCIA
                Set rsInc = Db.OpenResultset("select d.nom from incidencias i left join dependentes d on i.usuario=d.codi where i.id=" & incOrigen)
                If Not rsInc.EOF Then usuarioNom = empOrigen & "  " & rsInc("nom")
            
                Set rsCreador = Db.OpenResultset("select * from " & empDestino & ".dbo.dependentes where nom = 'CERBERO'")
                If Not rsCreador.EOF Then
                    creador = rsCreador("codi")
                Else
                    creador = "1"
                    Set rsCreador = Db.OpenResultset("select max(c) + 1 codi from (select max(codi) c from " & empDestino & ".dbo.dependentes union select max(codi) c from " & empDestino & ".dbo.dependentes_zombis) k")
                    If Not rsCreador.EOF Then creador = rsCreador("codi")

                    ExecutaComandaSql "insert into " & empDestino & ".dbo.dependentes (codi, nom, memo, [Hi Editem Horaris]) values(" & creador & ", 'CERBERO', 'CERBERO', 1)"
                End If
                
                If empDestino <> "" Then
                    sql = "INSERT INTO " & empDestino & ".dbo.incidencias (TimeStamp, Tipo, Usuario, Cliente, Estado, Prioridad, Tecnico, Observaciones, llamada, incidencia) "
                    sql = sql & "select CURRENT_TIMESTAMP, 'Incidència', '" & creador & "', '" & cliDestino & "', estado, prioridad, " & respDestino & ", Observaciones, llamada, incidencia "
                    sql = sql & "from incidencias where id = " & incOrigen
                    ExecutaComandaSql sql
    
                    Set rsInc = Db.OpenResultset("select top 1 id, incidencia from " & empDestino & ".dbo.incidencias where cliente='" & cliDestino & "' and usuario=" & creador & " and tecnico=" & respDestino & " order by timestamp desc", rdConcurRowVer)
                    If Not rsInc.EOF Then
                        idInc = rsInc("id")
                    
                        'ALTA EN EL HISTORICO
                        sql = "INSERT INTO " & empDestino & ".dbo.Inc_Historico (Id, Timestamp, usuario, incidencia, tipo) VALUES "
                        sql = sql & "(" & idInc & ", getdate(), '" & usuarioNom & "', '" & rsInc("incidencia") & "', 'TEXTO')"
                        ExecutaComandaSql sql
                
                        ExecutaComandaSql "insert into Inc_Link_Otros (id, idOtro, Empresa) values (" & incOrigen & ", " & idInc & ", '" & empDestino & "')"
                        ExecutaComandaSql "insert into " & empDestino & ".dbo.Inc_Link_Otros (id, idOtro, Empresa) values (" & idInc & ", " & incOrigen & ", '" & empOrigen & "')"
                        
                        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param3 = '" & idInc & "' where idCodigo='" & codigoDeAccion & "'"
                    End If
                End If
            Else 'HAY CORRESPONDENCIA CON OTRA TAREA
                Set rsInc = Db.OpenResultset("select * from incidencias where id=" & incOrigen)
                If Not rsInc.EOF Then
                    If InStr(items, "TEXTO") Then
                        Set rsHist = Db.OpenResultset("select * from Inc_Historico where id=" & incOrigen & " and tipo='TEXTO' and timestamp = '" & fechaHistorico & "'", rdConcurRowVer)
                        If Not rsHist.EOF Then
                            ExecutaComandaSql "INSERT INTO " & empDestino & ".dbo.Inc_Historico (Id, Timestamp, usuario, incidencia, tipo) values (" & incDestino & ", getdate(), '" & empOrigen & "-" & rsHist("usuario") & "', '" & rsHist("incidencia") & "', 'TEXTO')"
                        End If
                    End If
                    
                    If InStr(items, "ESTADO") Then
                        Set rsHist = Db.OpenResultset("select * from Inc_Historico where id=" & incOrigen & " and tipo='ESTADO' and timestamp = '" & fechaHistorico & "'", rdConcurRowVer)
                        If Not rsHist.EOF Then
                            ExecutaComandaSql "INSERT INTO " & empDestino & ".dbo.Inc_Historico (Id, Timestamp, usuario, incidencia, tipo) values (" & incDestino & ", getdate(), '" & empOrigen & "-" & rsHist("usuario") & "', '" & rsHist("incidencia") & "', 'ESTADO')"
                            
                            If rsHist("Incidencia") = "Cerrada" Then
                                ExecutaComandaSql "UPDATE " & empDestino & ".dbo.incidencias SET Estado = 'Cerrada', FFinReparacion = GETDATE(), lastUpdate = GETDATE() WHERE id = " & incDestino
                            Else
                                ExecutaComandaSql "UPDATE " & empDestino & ".dbo.incidencias SET Estado = '" & rsInc("estado") & "', FFinReparacion = NULL, lastUpdate = GETDATE() WHERE id = " & incDestino
                            End If
                        End If
                    End If
               
                    If InStr(items, "PRIORIDAD") Then
                        Set rsHist = Db.OpenResultset("select * from Inc_Historico where id=" & incOrigen & " and tipo='PRIORIDAD' and timestamp = '" & fechaHistorico & "'", rdConcurRowVer)
                        If Not rsHist.EOF Then
                            ExecutaComandaSql "INSERT INTO " & empDestino & ".dbo.Inc_Historico (Id, Timestamp, usuario, incidencia, tipo) values (" & incDestino & ", getdate(), '" & empOrigen & "-" & rsHist("usuario") & "', '" & rsHist("incidencia") & "', 'PRIORIDAD')"
                            
                            ExecutaComandaSql "UPDATE " & empDestino & ".dbo.incidencias set prioridad = '" & rsInc("prioridad") & "' where id=" & incDestino
                        End If
                    End If

                End If
            End If
            
        End If
        

    End If
    
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR InterpretaIncidenciaCERBERO", "ERROR: " & err.Description, "", ""

End Sub


Sub preparaCodigoDeAccionFichaje(HTMLText As String, emailDe As String, empresa As String)
    Dim HTMLAux As String
    Dim codigoDeAccion As String
    Dim turnoSeleccionado As String, turnosExtra As String, turnosAprendiz As String, turnosCoord As String, nuevoEntra As String, nuevoSale As String
    Dim entraOK As Boolean, saleOK As Boolean
    Dim tagCdA As String, tagTurno As String, tagTurnoExtra As String, tagAprendiz As String, tagCoordinacion As String
    Dim rsCodigoDeAccion As rdoResultset, rsDep As rdoResultset, tipoEmpleado As String, rsIdTurno As rdoResultset
    
    On Error GoTo nor:
    
    HTMLAux = UCase(HTMLText)
    
    tagCdA = "CODIGO_ACCION:["
    tagTurno = "NAME=""T_"
    tagTurnoExtra = "NAME=""T_EXTRA"
    tagAprendiz = "NAME=""T_APRENDIZ"
    tagCoordinacion = "NAME=""T_COORDINACION"
    
    If InStr(HTMLAux, tagCdA) Then
        codigoDeAccion = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA), 100)
        codigoDeAccion = Mid(codigoDeAccion, 1, InStr(codigoDeAccion, "]") - 1)
        Set rsCodigoDeAccion = Db.OpenResultset("select * from  " & taulaCodigosDeAccion() & "  where idCodigo='" & codigoDeAccion & "'")
        
        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param10 = '" & emailDe & "' where idCodigo='" & codigoDeAccion & "'"
        
        HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA))
        
        If InStr(HTMLAux, ">OK<") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, ">OK<"))
            'HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "NAME=""T_"))
            turnoSeleccionado = Mid(HTMLAux, InStr(HTMLAux, tagTurno) + Len(tagTurno), 100)
            turnoSeleccionado = Mid(turnoSeleccionado, 1, InStr(turnoSeleccionado, """") - 1)
            
            Set rsIdTurno = Db.OpenResultset("select * from cdpTurnos where idTurno='" & turnoSeleccionado & "'")
            If Not rsIdTurno.EOF Then ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param5 = '" & turnoSeleccionado & "' where idCodigo='" & codigoDeAccion & "'"
            
            'Mirar si hay turno + extras
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagTurnoExtra))
            turnosExtra = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1)
            'If turnosExtra = "-" Then turnosExtra = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1) 'BUSCAR EL SIGUIENTE < Y COGER HASTA AHÍ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If IsNumeric(turnosExtra) Then
                ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param6 = '" & turnosExtra & "' where idCodigo='" & codigoDeAccion & "'"
            End If
        End If
        
        If InStr(HTMLAux, "ENTRADANUEVO") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "ENTRADANUEVO"))
            nuevoEntra = Trim(Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1))
        End If
        
        If InStr(HTMLAux, "SALIDANUEVO") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "SALIDANUEVO"))
            nuevoSale = Trim(Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1))
        End If
        
        'GENERACIÓN DE TURNO NUEVO
        entraOK = False
        If IsNumeric(nuevoEntra) Then  'Hora sin minutos
            If CInt(nuevoEntra) < 24 Then
                nuevoEntra = Right("00" & nuevoEntra, 2) & ":00"
                entraOK = True
            End If
        Else
            If InStr(nuevoEntra, ":") Then 'Hora con minutos
                If IsNumeric(Split(nuevoEntra, ":")(0)) And IsNumeric(Split(nuevoEntra, ":")(1)) Then
                    If CInt(Split(nuevoEntra, ":")(0)) < 24 And CInt(Split(nuevoEntra, ":")(1)) < 59 Then
                        nuevoEntra = Right("00" & Trim(Split(nuevoEntra, ":")(0)), 2) & ":" & Right("00" & Trim(Split(nuevoEntra, ":")(1)), 2)
                        entraOK = True
                    End If
                End If
            End If
        End If

        saleOK = False
        If IsNumeric(nuevoSale) Then  'Hora sin minutos
            If CInt(nuevoSale) < 24 Then
                nuevoSale = Right("00" & nuevoSale, 2) & ":00"
                saleOK = True
            End If
        Else
            If InStr(nuevoSale, ":") Then 'Hora con minutos
                If IsNumeric(Split(nuevoSale, ":")(0)) And IsNumeric(Split(nuevoSale, ":")(1)) Then
                    If CInt(Split(nuevoSale, ":")(0)) < 24 And CInt(Split(nuevoSale, ":")(1)) < 59 Then
                        nuevoSale = Right("00" & Trim(Split(nuevoSale, ":")(0)), 2) & ":" & Right("00" & Trim(Split(nuevoSale, ":")(1)), 2)
                        saleOK = True
                    End If
                End If
            End If
        End If

        If entraOK And saleOK Then
            'BUSCAR TIPO DE TRABAJADOR QUE HA FICHADO
            tipoEmpleado = "DEPENDENTA"
            Set rsDep = Db.OpenResultset("select * from dependentesextes where nom ='TIPUSTREBALLADOR' and id=" & rsCodigoDeAccion("param2"))
            If Not rsDep.EOF Then tipoEmpleado = rsDep("valor")
            
            Set rsIdTurno = Db.OpenResultset("select * from cdpTurnos where horaInicio = '" & nuevoEntra & "' and horaFin = '" & nuevoSale & "' and tipoEmpleado like '%" & tipoEmpleado & "%'")
            If Not rsIdTurno.EOF Then
                turnoSeleccionado = rsIdTurno("idTurno")
            Else
                Set rsIdTurno = Db.OpenResultset("select newid() Id")
                turnoSeleccionado = rsIdTurno("id")
                
                ExecutaComandaSql "insert into cdpTurnos (nombre , horaInicio, horaFin, idTurno, color, tipoEmpleado) values ('De " & nuevoEntra & " a " & nuevoSale & "', '" & nuevoEntra & "', '" & nuevoSale & "', '" & turnoSeleccionado & "', '#DDDDDD', '" & tipoEmpleado & "')"
            End If
            
            ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param5 = '" & turnoSeleccionado & "' where idCodigo='" & codigoDeAccion & "'"
        End If
        
        If InStr(HTMLAux, "HORASEXTRA") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "HORASEXTRA"))
            turnosExtra = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1)
            'If turnosExtra = "-" Then turnosExtra = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, 2)
            If IsNumeric(turnosExtra) Then
                ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param6 = '" & turnosExtra & "' where idCodigo='" & codigoDeAccion & "'"
            End If
        End If
        
        If InStr(HTMLAux, "HORASAPRENDIZ") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "HORASAPRENDIZ"))
            turnosAprendiz = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1)
            If IsNumeric(turnosAprendiz) Then
                ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param7 = '" & turnosAprendiz & "' where idCodigo='" & codigoDeAccion & "'"
            End If
        End If
        
        If InStr(HTMLAux, "HORASCOORDINACION") Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, "HORASCOORDINACION"))
            turnosCoord = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1)
            If IsNumeric(turnosCoord) Then
                ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param8 = '" & turnosCoord & "' where idCodigo='" & codigoDeAccion & "'"
            End If
        End If
        
        
        ExecutaComandaSql "insert into feinesafer (Id, Tipus, Ciclica, Param1, tmStmp) values (newid(), 'CodigoDeAccion', 0, '" & codigoDeAccion & "', getdate())"
    End If
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR preparaCodigoDeAccionFichaje", "ERROR: " & err.Description, "", ""
End Sub


Sub preparaCodigoDeAccionCuadrante(HTMLText As String, emailDe As String, empresa As String)
    Dim sql As String
    Dim HTMLAux As String
    Dim codigoDeAccion As String, botiga As String
    Dim nuevoEntra As String, nuevoSale As String, turnoSeleccionado As String
    Dim entraOK As Boolean, saleOK As Boolean
    Dim tagCdA As String, tagCuadrante As String, tagTipoEmpleado As String, tagEntrada As String, tagSalida As String, tagFecha As String
    Dim rsCodigoDeAccion As rdoResultset, rsDep As rdoResultset, tipoEmpleado As String, rsIdTurno As rdoResultset
    Dim fecha As Date, periode As String, fechaStr As String
    Dim tagValidacion As String, tagValidado As String, validadoParams() As String
    
    On Error GoTo nor:
    
    HTMLAux = UCase(Replace(HTMLText, vbCrLf, ""))
    HTMLAux = Replace(HTMLAux, "X_", "")
    
    tagCdA = "CODIGO_ACCION:["
    tagValidacion = "NAME=""VALIDACION_TURNOS"
    tagValidado = "NAME=""VALIDADO_"
    tagCuadrante = "NAME=""CUADRANTE"
    tagTipoEmpleado = "NAME=""TIPOEMPLEADO_"
    tagEntrada = "NAME=""ENTRADA_"
    tagSalida = "NAME=""SALIDA_"
    tagFecha = "NAME=""FECHA_"
    
     
    If InStr(HTMLAux, tagCdA) Then
        codigoDeAccion = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA), 100)
        codigoDeAccion = Mid(codigoDeAccion, 1, InStr(codigoDeAccion, "]") - 1)
        Set rsCodigoDeAccion = Db.OpenResultset("select * from  " & taulaCodigosDeAccion() & "  where idCodigo='" & codigoDeAccion & "'")
        If Not rsCodigoDeAccion.EOF Then botiga = rsCodigoDeAccion("param1")
        
        ExecutaComandaSql "update  " & taulaCodigosDeAccion() & "  set param10 = '" & emailDe & "' where idCodigo='" & codigoDeAccion & "'"
        
        If botiga <> "" Then
            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagCdA) + Len(tagCdA))
            
            'Validación turnos
            If InStr(HTMLAux, tagValidacion) Then
                HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagValidacion))
                While InStr(HTMLAux, tagValidado)
                    HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagValidado))
                    validadoParams = Split(Mid(HTMLAux, Len("NAME=""") + 1, InStr(HTMLAux, ">") - Len("NAME=""") - 2), "_")
            
                    ExecutaComandaSql "delete from " & taulaCdpValidacionHoras(CDate(validadoParams(1))) & " where botiga='" & validadoParams(2) & "' and fecha = '" & validadoParams(1) & "'"
                    If InStr(HTMLAux, ">OK<") Then
                        sql = "insert into " & taulaCdpValidacionHoras(CDate(validadoParams(1))) & " (idPlan, fecha, botiga, dependenta, usuarioModif, validado) "
                        sql = sql & "select idPlan, '" & validadoParams(1) & "', '" & botiga & "', idEmpleado, '" & emailDe & "', 1 "
                        sql = sql & "from " & taulaCdpPlanificacion(CDate(validadoParams(1))) & " "
                        sql = sql & "where botiga=" & botiga & " and day(fecha)=" & Day(CDate(validadoParams(1))) & " and activo=1"
                        
                        ExecutaComandaSql sql
                        'ExecutaComandaSql "insert into " & taulaCdpValidacionHoras(CDate(validadoParams(1))) & " (idPlan, fecha, botiga, dependenta, usuarioModif, validado) values ('" & validadoParams(1) & "', '" & validadoParams(2) & "', '" & emailDe & "', 1)"
                    End If
            
                    HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagValidado) + Len(tagValidado))
                Wend
            
            End If
            
            'Modificación cuadrante
            If InStr(HTMLAux, tagCuadrante) Then
                HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagCuadrante))
    
                While InStr(HTMLAux, tagFecha)
                    HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagFecha))
                    fechaStr = Mid(HTMLAux, Len("NAME=""") + 1, InStr(HTMLAux, ">") - Len("NAME=""") - 2)
                    
                    fecha = CDate(Split(fechaStr, "_")(1))
                    If fecha > Now() Then
                        'Eliminamos turnos definidos para ese día, para poder poner los nuevos
                        ExecutaComandaSql "update " & taulaCdpPlanificacion(fecha) & " set activo=0 where botiga=" & botiga & " and day(fecha)=" & Day(fecha)
                        'ExecutaComandaSql "delete from " & taulaCdpPlanificacion(fecha) & " where botiga=" & botiga & " and day(fecha)=" & Day(fecha)
                    
                        While InStr(HTMLAux, tagTipoEmpleado & Format(fecha, "dd/mm/yyyy"))
                            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagTipoEmpleado & Format(fecha, "dd/mm/yyyy")))
                            tipoEmpleado = Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1)
                    
                            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagEntrada))
                            nuevoEntra = Trim(Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1))
                            
                            HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagSalida))
                            nuevoSale = Trim(Mid(HTMLAux, InStr(HTMLAux, ">") + 1, (InStr(HTMLAux, "<") - 2) - InStr(HTMLAux, ">") + 1))
                                            
                            'GENERACIÓN DE TURNO NUEVO
                            entraOK = False
                            If IsNumeric(nuevoEntra) Then  'Hora sin minutos
                                If CInt(nuevoEntra) < 24 Then
                                    nuevoEntra = Right("00" & nuevoEntra, 2) & ":00"
                                    entraOK = True
                                End If
                            Else
                                If InStr(nuevoEntra, ":") Then 'Hora con minutos
                                    If IsNumeric(Split(nuevoEntra, ":")(0)) And IsNumeric(Split(nuevoEntra, ":")(1)) Then
                                        If CInt(Split(nuevoEntra, ":")(0)) < 24 And CInt(Split(nuevoEntra, ":")(1)) < 59 Then
                                            nuevoEntra = Right("00" & Trim(Split(nuevoEntra, ":")(0)), 2) & ":" & Right("00" & Trim(Split(nuevoEntra, ":")(1)), 2)
                                            entraOK = True
                                        End If
                                    End If
                                End If
                            End If
                    
                            saleOK = False
                            If IsNumeric(nuevoSale) Then  'Hora sin minutos
                                If CInt(nuevoSale) < 24 Then
                                    nuevoSale = Right("00" & nuevoSale, 2) & ":00"
                                    saleOK = True
                                End If
                            Else
                                If InStr(nuevoSale, ":") Then 'Hora con minutos
                                    If IsNumeric(Split(nuevoSale, ":")(0)) And IsNumeric(Split(nuevoSale, ":")(1)) Then
                                        If CInt(Split(nuevoSale, ":")(0)) < 24 And CInt(Split(nuevoSale, ":")(1)) < 59 Then
                                            nuevoSale = Right("00" & Trim(Split(nuevoSale, ":")(0)), 2) & ":" & Right("00" & Trim(Split(nuevoSale, ":")(1)), 2)
                                            saleOK = True
                                        End If
                                    End If
                                End If
                            End If
                    
                            If entraOK And saleOK Then
                                If tipoEmpleado <> "D" And tipoEmpleado <> "F" Then tipoEmpleado = "RESPONSABLE/DEPENDENTA"
                                If tipoEmpleado = "D" Then tipoEmpleado = "RESPONSABLE/DEPENDENTA"
                                If tipoEmpleado = "F" Then
                                tipoEmpleado = "FORNER"
                                End If
                                
                                periode = "M"
                                If Split(nuevoEntra, ":")(0) >= 14 Then periode = "T"
                                
                                Set rsIdTurno = Db.OpenResultset("select * from cdpTurnos where horaInicio = '" & nuevoEntra & "' and horaFin = '" & nuevoSale & "' and tipoEmpleado like '%" & tipoEmpleado & "%'")
                                If Not rsIdTurno.EOF Then
                                    turnoSeleccionado = rsIdTurno("idTurno")
                                Else
                                    Set rsIdTurno = Db.OpenResultset("select newid() Id")
                                    turnoSeleccionado = rsIdTurno("id")
                                    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "INSERTA TURNO NUEVO ", "insert into cdpTurnos (nombre , horaInicio, horaFin, idTurno, color, tipoEmpleado) values ('De " & nuevoEntra & " a " & nuevoSale & "', '" & nuevoEntra & "', '" & nuevoSale & "', '" & turnoSeleccionado & "', '#DDDDDD', '" & tipoEmpleado & "')", "", ""
                                    ExecutaComandaSql "insert into cdpTurnos (nombre , horaInicio, horaFin, idTurno, color, tipoEmpleado) values ('De " & nuevoEntra & " a " & nuevoSale & "', '" & nuevoEntra & "', '" & nuevoSale & "', '" & turnoSeleccionado & "', '#DDDDDD', '" & tipoEmpleado & "')"
                                End If
                            
                                sql = "insert into " & taulaCdpPlanificacion(fecha) & " (idPlan, fecha, botiga, periode, idTurno, usuarioModif, fechaModif, activo) values "
                                sql = sql & "(newid(), '" & fecha & "', '" & botiga & "', '" & periode & "', '" & turnoSeleccionado & "', '" & emailDe & "', getdate(), 1)"
                                ExecutaComandaSql sql
                            End If
                            
                        Wend
                        
                    Else
                        HTMLAux = Mid(HTMLAux, InStr(HTMLAux, tagFecha) + Len(tagFecha))
                    End If
                Wend
            End If
        End If
    End If
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR preparaCodigoDeAccionCuadrante", "ERROR: " & err.Description & "  SQL: " & sql, "", ""
End Sub


Sub revisaEmail()

    'RevisaEmailDe "secrehit@gmail.com", "secrehit2130"
    'RevisaEmailDe "secreSilema@gmail.com", "LOperas93786"
    ' CONTRASEÑA APLICACION
'    RevisaEmailDe "Sandra@HitSystems.es", "elefante1234", "Xml"
    RevisaEmailDe "secrehit@gmail.com", "1234secrehit"
    RevisaEmailDe "SecreHit@hit.cat", "secrehit1234"
    'RevisaEmailDe "email@hit.cat", "emailhit"

End Sub

'Sub RevisaEmailDe(User As String, Password As String)
'    Dim UnDeBorrat As Boolean, NomAtt As String, H1 As Date, i, downHTTP As New HTTP, Empresa As String, Subj, t, Id, nomF, ruta1, ruta2, ruta3, sql, extF, mimeF, NomUsu, n, rs, URL, resposta, aResposta, emailDe As String, IdFoto As String
'On Error GoTo nor
'    InformaEmpresa "Hit Email " & User
'    Set frmSplash.Pop3 = New wodPop3Com
'    frmSplash.Pop3.HostName = "pop.gmail.com"
'    frmSplash.Pop3.Security = SecurityImplicit
'
'
'
''ExecutaComandaSql "use fac_papa"
''EnviaEmailAdjunto "jordi@hit.cat", "Excel Resultat ", EnviarResultatBotiga(1)
'
'
'    frmSplash.Pop3.Port = 995
'    frmSplash.Pop3.Login = User
'    frmSplash.Pop3.Password = Password
'    frmSplash.Pop3.Connect
'
'    Informa "Conectant " & frmSplash.Pop3.Login
'    H1 = Now
'    While Not frmSplash.Pop3.State =  WODPOP3COMLib.StatesEnum.Connected
'        DoEvents
'        If Now > DateAdd("s", 90, H1) Then Exit Sub
'    Wend
'    Informa "Conectat " & frmSplash.Pop3.Messages.Count & " Emails"
'
'    For i = 0 To frmSplash.Pop3.Messages.Count - 1
'        Set frmSplash.Pop3Message = Nothing
'        Informa "Llegint " & i & " De " & frmSplash.Pop3.Messages.Count & " "
'        frmSplash.Pop3.Messages(i).Get
'        H1 = Now
'        While frmSplash.Pop3Message Is Nothing
'            DoEvents
'            If Now > DateAdd("s", 90, H1) Then Exit Sub
'        Wend
'
'
'        Informa frmSplash.Pop3Message.Subject & "(" & frmSplash.Pop3Message.Attachments.Count & ")"
'        emailDe = frmSplash.Pop3Message.FromEmail
'        Empresa = SecreHitEmailEmpresa(emailDe)
'        If Empresa = "" Then
'             EnviaEmail emailDe, "No Te Conozco :( , i no hablo con desconocidos..."
'        Else
'           If frmSplash.Pop3Message.Attachments.Count = 0 Then
'                Subj = frmSplash.Pop3Message.Subject
'                If UCase(Left(Subj, 3)) = UCase("Tel") Then
'                   NomUsu = Trim(Right(Subj, Len(Subj) - InStr(Subj, " ")))
'                   Dim TelRandom As String
'                   TelRandom = Right("3" & (Rnd * 10000000), 6)
'                   ExecutaComandaSql "Delete hit.dbo.Telefonos where idt = '" & NomUsu & "'"
'                   ExecutaComandaSql "Insert Into hit.dbo.Telefonos (idt,cliente,actdata,actcodi,actestat,domini) Values ('" & NomUsu & "','" & Empresa & "',GETDATE(),'" & TelRandom & "','PendentActivacio','') "
'                   ExecutaComandaSql "Insert Into " & Empresa & ".dbo.Recurssos (idt,cliente,actdata,actcodi,actestat,domini) Values ('" & NomUsu & "','" & Empresa & "',GETDATE(),'" & TelRandom & "','PendentActivacio','') "
'                   sf_enviarMail "Secrehit@gmail.com", emailDe, "Telefono : " & NomUsu & " Asignado.", "Desde el telefon " & NomUsu & " llame al 937160210 i cuando se lo pian teclee el codigo : <Br> " & TelRandom & Chr(13) & Chr(10) & " <Br><Br><Br> Atentamente Joana.", "", ""
'
'                End If
'                If UCase(Left(Subj, 9)) = UCase("informe v") Then
'                    ExecutaComandaSql "use  " & Empresa
'                    Set rs = Db.OpenResultset("Select distinct Botiga from [" & NomTaulaVentas(Now) & "] where day(data) = " & Day(Now))
'                    While Not rs.EOF
'                        EnviarReportBotiga CDbl(rs("Botiga")), emailDe
'                        rs.MoveNext
'                    Wend
'                    ExecutaComandaSql "use Hit"
'                End If
'                If UCase(Left(Subj, 8)) = UCase("Resultat") Then
'                    ExecutaComandaSql "use  " & Empresa
'                    Set rs = Db.OpenResultset("Select distinct Botiga from [" & NomTaulaVentas(Now) & "] where day(data) = " & Day(Now))
'                    While Not rs.EOF
'                        EnviaEmailAdjunto emailDe, "Excel Resultat ", EnviarResultatBotiga(rs("Botiga"))
'                        rs.MoveNext
'                    Wend
'                    ExecutaComandaSql "use Hit"
'                End If
'
'           Else
'                extF = "Jpg"
'                If InStr(frmSplash.Pop3Message.Attachments(0).Name, ".") > 0 Then extF = Right(frmSplash.Pop3Message.Attachments(0).Name, Len(frmSplash.Pop3Message.Attachments(0).Name) - InStr(frmSplash.Pop3Message.Attachments(0).Name, "."))
'                mimeF = "image/jpeg"
'                NomAtt = "AttEmail." & extF
'                MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
'                frmSplash.Pop3Message.Attachments(0).Save "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
'                Informa "Rebut File " & frmSplash.Pop3Message.Attachments(0).Name
'                t = ""
'                Id = ""
'                Subj = frmSplash.Pop3Message.Subject
'                If UCase(Left(Subj, 3)) = "ART" Then
'                    t = "articles"
'                    Set rs = Db.OpenResultset("Select codi,Nom From  " & Empresa & ".dbo.articles where nom like '%" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))) & "%' ")
'                    If rs.EOF Then Set rs = Db.OpenResultset("select a.codi codi , a.nom nom  from  " & Empresa & ".dbo.articles a join  " & Empresa & ".dbo.articlespropietats p on a.codi = p.codiarticle where variable = 'CODI_PROD'  and valor = '" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))) & "' ")
'
'                    If Not rs.EOF Then
'                        Id = rs(0)
'                        NomUsu = rs("nom")
'                    End If
'                End If
'
'                If UCase(Left(Subj, 3)) = "TRE" Or UCase(Left(Subj, 3)) = "DEP" Then
'                    t = "dependentes"
'                    Dim NomDep As String
'                    NomDep = Trim(Right(Subj, Len(Subj) - InStr(Subj, " ")))
'                    Set rs = Db.OpenResultset("Select codi,nom From " & Empresa & ".dbo.dependentes where nom like '%" & NomDep & "%' ")
'                    If Not rs.EOF Then
'                        Id = rs("Codi")
'                        NomUsu = rs("nom")
'                    Else
'                       ExecutaComandaSql "insert into " & Empresa & ".dbo.dependentes  Select max(codi)+1 as codi ,'" & NomDep & "' as Nom , '" & NomDep & "' as memo, '' as telefon , '' as [adreça] , '' as icona, 0 as [hi editem horaris] , '' tid from " & Empresa & ".dbo.dependentes "
'                       Set rs = Db.OpenResultset("Select codi,nom From " & Empresa & ".dbo.dependentes where nom like '%" & NomDep & "%' ")
'                       If Not rs.EOF Then
'                          Id = rs("Codi")
'                          NomUsu = rs("nom")
'                       End If
'                    End If
'                End If
'
'                If UCase(Left(Trim(Subj), 3)) = "FAC" Then
'                    t = "Factura"
'                    InterpretaFacturaPdf "\Facturacion\ElForn\file\tmp\" & NomAtt, Empresa, emailDe
'                End If
'
'                If UCase(Left(Subj, 3)) = "INC" Then
'                    t = "Incidencias"
'                    If InStr(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), " ") > 0 Then
'                        If IsNumeric(Left(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), InStr(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), " "))) Then
'                            Set rs = Db.OpenResultset("Select id From " & Empresa & ".dbo.incidencias where id = '" & Left(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), InStr(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), " ")) & "' ")
'                        Else
'                            Set rs = Db.OpenResultset("Select id From " & Empresa & ".dbo.incidencias where id = '-9999' ")
'                        End If
'
'                    Else
'                        If IsNumeric(Trim(Right(Subj, Len(Subj) - InStr(Subj, " ")))) Then
'                            Set rs = Db.OpenResultset("Select id From " & Empresa & ".dbo.incidencias where id = '" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))) & "' ")
'                        Else
'                            Set rs = Db.OpenResultset("Select id From " & Empresa & ".dbo.incidencias where id = '-9999' ")
'                        End If
'                    End If
'                    If Not rs.EOF Then
'                        ruta1 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
'                        If InStr(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), " ") > 0 Then
'                            salvar ruta1, "INC_" & Left(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), InStr(Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), " ")), extF, mimeF, "INC_" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), NomUsu, n, t, Id, Empresa
'                        Else
'                            salvar ruta1, "INC_" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), extF, mimeF, "INC_" & Trim(Right(Subj, Len(Subj) - InStr(Subj, " "))), NomUsu, n, t, Id, Empresa
'                        End If
'                        MyKill CStr(ruta1)
'                        EnviaEmail emailDe, "Foto de incidencia " & rs("id") & " Actualizada"
'                    Else
'                        EnviaEmail emailDe, "Error recibiendo mail no encuentro incidencia, (" & Subj & "). Formato INC [ESPACIO] NUMERO DE INCIDENCIA"
'                    End If
'                End If
'
'
'                If t <> "" And Id <> "" Then
'                    URL = "http://www.gestiondelatienda.com/facturacion/elforn/file/tmp/uploadFoto.php?"
'                    URL = URL & "rutaImg=" & NomAtt
'                    resposta = llegeigHtml(URL)
'                    aResposta = Split(resposta, "|")
'                    If UBound(aResposta) > 0 Then
'                        'Foto
'                        ruta1 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(0)
'                        salvar ruta1, "ORIGINAL", extF, mimeF, "Foto original de " & NomUsu, NomUsu, n, t, Id, Empresa
'                        MyKill CStr(ruta1)
'                        'Foto Screen
'                        ruta2 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(1)
'                        salvar ruta2, "SCREEN", extF, mimeF, "Foto pantalla de " & NomUsu, NomUsu, n, t, Id, Empresa
'                        MyKill CStr(ruta2)
'                        'ICO
'                        ruta3 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(2)
'                        IdFoto = salvar(ruta3, "ICO", extF, mimeF, "Foto TPV de " & NomUsu, NomUsu, n, t, Id, Empresa)
'                        'Foto TPV
'                        IdFoto = salvar(ruta1, "TPV", extF, mimeF, "Foto TPV de " & NomUsu, NomUsu, n, t, Id, Empresa)
'                        MyKill CStr(ruta3)
'                        EnviaEmail emailDe, "Foto de " & NomUsu & " Actualizada"
'                    Else
'                        EnviaEmail emailDe, "Error de imatge " & resposta & " ho sento :("
'                    End If
'                    'Pujar imatge  a botigues
'                    rec ("INSERT into missatgesaenviar (Tipus,Param) values ('Imatges" & t & "','') ")
'                Else
'                    If t <> "Incidencias" And t <> "Factura" Then
'                        EnviaEmail emailDe, "No Se para quien es la foto, (" & Subj & "). Formato ART i nom o codi article o DEP i part del nom de la dependenta"
'                    End If
'                End If
'                MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
'            End If
'        End If
'        frmSplash.Pop3.Messages(i).MarkDelete = True
'        UnDeBorrat = True
'    Next
'
'    frmSplash.Pop3.Disconnect
'    H1 = Now
'    While frmSplash.Pop3.State =  WODPOP3COMLib.StatesEnum.Connected
'        DoEvents
'        If Now > DateAdd("s", 90, H1) Then Exit Sub
'    Wend
'
'    If UnDeBorrat Then
'        H1 = Now
'        While Now <= DateAdd("s", 4, H1)
'            DoEvents
'        Wend
'    End If
'
'nor:
'End Sub



Function salvar(ByVal ruta, ByVal nom, ByVal ext, ByVal mime, ByVal desc, ByVal usu, ByVal n, t, iD, empresa) As String
    Dim rs As ADODB.Recordset, st, sql, rsId, IdFoto, FS, Tipus, rsIns, rsUpd, Rs2 As rdoResultset
    'Dim oFile As New Scripting.FileSystemObject
    Dim FileExists As Boolean
    FileExists = False
    On Error Resume Next
        db2.Close
    On Error GoTo 0
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & empresa & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"
    Set rs = rec("select top 1 * from " & empresa & ".dbo." & tablaArchivo(), True)
    rs.AddNew
    
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    'FileExists = oFile.FileExists(ruta)
    If Dir$(ruta) <> "" Then
        FileExists = True
    End If
    If FileExists Then
        st.LoadFromFile ruta
    Else
        ruta = "C:\Web\gdt\Facturacion\ElForn\file\tmp\AttEmail.jpg"
     st.LoadFromFile ruta
    End If
    Set Rs2 = Db.OpenResultset("select newid() i")
    IdFoto = Rs2(0)
    
    rs.Fields("id").Value = IdFoto
    rs.Fields("nombre").Value = nom
    rs.Fields("archivo").Value = st.Read
    rs.Fields("extension").Value = ext
    rs.Fields("mime").Value = mime
    rs.Fields("descripcion").Value = desc
    rs.Fields("fecha").Value = Now
    rs.Fields("propietario").Value = usu
    rs.Fields("tmp").Value = 0
    rs.Fields("down").Value = 0
    st.Close
    rs.Update
    rs.Close
    'Update taules extes per lligar fitxer amb usuari/recurs/article
    Select Case nom
        Case "ORIGINAL"
            Tipus = "FOTO"
        Case "TPV"
            Tipus = "FOTOTPV"
        Case "SCREEN"
            Tipus = "FOTOSCREEN"
    End Select
    If Left(nom, 3) <> "INC" Then
        sql = "select valor from  " & empresa & ".dbo." & t & "Extes where id='" & iD & "' "
        sql = sql & " and nom = '" & Tipus & IIf(n > 0, n, "") & "'"
        Set rs = rec(sql)
        If rs.EOF Then
            sql = "INSERT into  " & empresa & ".dbo." & t & "Extes values ('" & iD & "','" & Tipus & IIf(n > 0, n, "") & "','" & IdFoto & "')"
            Set rsIns = rec(sql)
        Else
            sql = "UPDATE  " & empresa & ".dbo." & t & "Extes set valor='" & IdFoto & "' where id='" & iD & "' and nom ='" & Tipus & IIf(n > 0, n, "") & "'"
            Set rsUpd = rec(sql)
        End If
        salvar = IdFoto
    End If
End Function


Sub InterpretaFacturaXml(File As String, empresa As String, emailDe As String, nombreFichero As String)
    Dim intFile As Integer
    Dim XMLDoc As DOMDocument, xNode As IXMLDOMNode
    Dim fNode As IXMLDOMNode, iNode As IXMLDOMNode, tNode As IXMLDOMNode, fNodeChild As IXMLDOMNode
    Set XMLDoc = New DOMDocument
    Dim lineXML As String, strXML As String
    Dim rs As rdoResultset, rsPedido As rdoResultset, rsFactura As rdoResultset, rsCliente As rdoResultset, rsXML As rdoResultset, sql As String, idFactura As String
    Dim fFactura As Date, fVencimiento As Date, nFactura As String, nPedido As String, idProveedor As String, cliente As String, fPedido As Date, codiArticle As String, nFacturaCorrectiva As String, fFacturaCorrectiva As Date
    Dim Total As Double, repartirImp As Double, llistaArt As String, llistaArtArr() As String, a As Integer, impTotal As Double, pctReparto As Double, preuArticle As Double
    Dim agrupar As Boolean, rsAgrupar As rdoResultset
    Dim nomArt As String, rsNomArt As rdoResultset

    On Error GoTo ERR_XML
    
    emailFactura_XML = ""
    
    nFacturaCorrectiva = ""
    
    
    'TABLA TEMPORAL PARA GUARDAR LOS LOTES DURANTE EL PROCESO
    ExecutaComandaSql "drop table " & empresa & ".dbo.[INTERPRETA_XML_TMP] "
    ExecutaComandaSql "CREATE TABLE " & empresa & ".dbo.[INTERPRETA_XML_TMP] ([idFactura] nvarchar(255) NULL, [idProducto] nvarchar(255) NULL, [nLot] nvarchar(255) NULL) ON [PRIMARY]"
    
    File = "C:\Web\gdt\Facturacion\ElForn\file\tmp\AttEmail.xml"
    
    'Open file
    intFile = FreeFile
    Open File For Input As intFile

    'Load XML into string strXML
    While Not EOF(intFile)
        Line Input #intFile, lineXML
        lineXML = Replace(lineXML, "ï»¿", "")
        strXML = strXML & lineXML
    Wend
    Close intFile

    XMLDoc.loadXML strXML

    Set rs = Db.OpenResultset("select newid() i")
    idFactura = rs("i")

    'Buscamos fecha y número factura
    For Each xNode In XMLDoc.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        If xNode.nodeName = "m:Facturae" Or xNode.nodeName = "fe:Facturae" Then
            For Each fNode In xNode.childNodes
                If fNode.nodeName = "Invoices" Then
                    InformaMiss "INTERPRETANT HEADER"
                    InterpretaFacturaXml_Header fNode.FirstChild, fFactura, fVencimiento, nFactura, nFacturaCorrectiva, fFacturaCorrectiva    'Fecha factura, Número factura
                    emailFactura_XML = emailFactura_XML & "<B>FECHA FACTURA: </B>" & fFactura & "<BR>"
                    emailFactura_XML = emailFactura_XML & "<B>FECHA VENCIMIENTO: </B>" & fVencimiento & "<BR>"
                    emailFactura_XML = emailFactura_XML & "<B>NUM FACTURA: </B>" & nFactura & "<BR>"
                End If
            Next
        End If
    Next
    
    'Si no hay empresa intentamos buscarla por Nif en todas nuestras bases de datos
    If empresa = "" Then
        For Each xNode In XMLDoc.childNodes
            InformaMiss "Nodo: " & xNode.nodeName, True
            If xNode.nodeName = "m:Facturae" Or xNode.nodeName = "fe:Facturae" Then
                For Each fNode In xNode.childNodes
                    Select Case fNode.nodeName
                        Case "Parties"
                            For Each fNodeChild In fNode.childNodes
                                Select Case fNodeChild.nodeName
                                    Case "BuyerParty"
                                        InformaMiss "BUSCANDO PROVEEDOR"
                                        InterpretaFacturaXml_Empresa fNodeChild, empresa
                                End Select
                            Next
                    End Select
                Next
            End If
        Next
    End If
    
    If empresa = "" Then GoTo ERR_XML_EMP
    
    'Buscamos proveedor
    For Each xNode In XMLDoc.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        If xNode.nodeName = "m:Facturae" Or xNode.nodeName = "fe:Facturae" Then
            For Each fNode In xNode.childNodes
                Select Case fNode.nodeName
                    Case "Parties"
                        For Each fNodeChild In fNode.childNodes
                            Select Case fNodeChild.nodeName
                                Case "SellerParty"
                                    InformaMiss "BUSCANDO PROVEEDOR"
                                    InterpretaFacturaXml_Proveedor fNodeChild, idProveedor, empresa
                            End Select
                        Next
                End Select
            Next
        End If
    Next
    
    If nFactura <> "" Then
        Set rsFactura = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] where numFactura = '" & nFactura & "' and EmpresaCodi = '" & idProveedor & "'")
        If rsFactura.EOF Then
            sql = "insert into [WEB]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] (IdFactura, NumFactura, DataInici, DataFi, DataFactura, DataEmissio, DataVenciment, BaseIva1, Iva1, BaseIva2, Iva2, BaseIva3, Iva3, BaseIva4, Iva4, BaseRec1, Rec1, BaseRec2, Rec2, BaseRec3, Rec3, BaseRec4, Rec4, "
            sql = sql & "valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, IvaRec1, IvaRec2, IvaRec3, IvaRec4, Reservat) "
            sql = sql & "values ('" & idFactura & "', '" & nFactura & "', convert(datetime,'" & fFactura & "',103), convert(datetime,'" & fFactura & "',103), convert(datetime,'" & fFactura & "',103), convert(datetime,'" & fFactura & "',103), convert(datetime,'" & fVencimiento & "',103), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 10, 21, 0, 0.5, 1.4, 4, 0, 0, 0, 0, 0, '')"
            ExecutaComandaSql sql
            
            agrupar = False
            Set rsAgrupar = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.ccProveedoresExtes where nom = 'FacturasAgrupadas' and valor='on' and id='" & idProveedor & "'")
            If Not rsAgrupar.EOF Then agrupar = True
            
            For Each xNode In XMLDoc.childNodes
                InformaMiss "Nodo: " & xNode.nodeName, True
                If xNode.nodeName = "m:Facturae" Or xNode.nodeName = "fe:Facturae" Then
                    For Each fNode In xNode.childNodes
                        Select Case fNode.nodeName
                            Case "Parties"
                                For Each fNodeChild In fNode.childNodes
                                    Select Case fNodeChild.nodeName
                                        Case "SellerParty"
                                            InformaMiss "INTERPRETANT SELLERPARTY"
                                            InterpretaFacturaXml_Seller fNodeChild, idFactura, fFactura, empresa  'EmpresaCodi, EmpNif, EmpNom, EmpAdresa, EmpCp, EmpCiutat, EmpTel, EmpFax, EmpeMail
                                        Case "BuyerParty"
                                            InformaMiss "INTERPRETANT BUYERPARTY"
                                            InterpretaFacturaXml_Buyer fNodeChild, idFactura, fFactura, empresa  'ClientCodi, ClientCodiFac, ClientNif, ClientNom, ClientAdresa, ClientCp, ClientCiutat, Tel, Fax, eMail
                                    End Select
                                Next
                        End Select
                    Next
                End If
            Next
            
            For Each xNode In XMLDoc.childNodes
                InformaMiss "Nodo: " & xNode.nodeName, True
                If xNode.nodeName = "m:Facturae" Or xNode.nodeName = "fe:Facturae" Then
                    For Each fNode In xNode.childNodes
                        If fNode.nodeName = "Invoices" Then
                            For Each iNode In fNode.FirstChild.childNodes
                                Select Case iNode.nodeName
                                    Case "TaxesOutputs"
                                        InterpretaFacturaXml_Taxes iNode, idFactura, fFactura, empresa 'BaseIva1, Iva1, BaseIva2, Iva2, BaseIva3, Iva3
                                    Case "InvoiceTotals"
                                        For Each tNode In iNode.childNodes
                                            InformaMiss "Nodo: " & tNode.nodeName, True
                                            If tNode.nodeName = "InvoiceTotal" Then Total = CDbl(tNode.FirstChild.nodeValue)
                                        Next
                                        
                                        emailFactura_XML = emailFactura_XML & "<B>TOTAL: </B>" & Total & "<BR>"
                                        ExecutaComandaSql "update [web]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] set Total=" & Total & " where idFactura='" & idFactura & "'"
                                    Case "Items"
                                        If Not ExisteixTaula("ccFacturasData_XML") Then ExecutaComandaSql "SELECT top 1 * into [WEB]." & empresa & ".dbo.ccFacturasData_XML FROM [WEB]." & empresa & ".dbo." & tablaFacturaProformaData(fFactura)
                                        
                                        ExecutaComandaSql "Delete from [WEB]." & empresa & ".dbo.ccFacturasData_XML"
                                        For Each tNode In iNode.childNodes
                                            InformaMiss "Nodo: " & tNode.nodeName, True
                                            If tNode.nodeName = "InvoiceLine" Then
                                                InterpretaFacturaXml_Linea tNode, idFactura, nFactura, fFactura, idProveedor, nPedido, empresa 'IdFactura, Data, Producte, ProducteNom, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat
                                            End If
                                        Next
                                        'If UCase(empresa) = UCase("Fac_Tena") Then 'Agrupar por IVA
                                        If agrupar Then
                                            sql = "insert into [WEB]." & empresa & ".dbo." & tablaFacturaProformaData(fFactura) & " "
                                            sql = sql & "SELECT IdFactura, '" & fFactura & "', null, newid(), 'Article No Codificat', null, sum(preu), sum(import), 0, tipusiva, iva, rec, referencia, 1, 0 "
                                            sql = sql & "From [WEB]." & empresa & ".dbo.ccFacturasData_XML "
                                            sql = sql & "WHERE IdFactura='" & idFactura & "' "
                                            sql = sql & "group by idFactura, TipusIva, Iva, rec, referencia"
                                            ExecutaComandaSql sql
                                        Else
                                            If UCase(empresa) = UCase("Fac_Tena") Then
                                                ExecutaComandaSql "Insert into [web]." & empresa & ".dbo.FeinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3) values (newid(), 'SolicitarAutorizacion', 0, '" & idFactura & "', '" & fFactura & "', 'ana@hit.cat')"
                                                ExecutaComandaSql "Insert into [web]." & empresa & ".dbo.FeinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3) values (newid(), 'SolicitarAutorizacion', 0, '" & idFactura & "', '" & fFactura & "', 'atena@silemabcn.com')"
                                            End If
                                        End If
                                        
                                        'ENVIAR A MURANO
                                        ExecutaComandaSql "Insert Into [web]." & empresa & ".dbo.FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFacturaRebuda', 0, '[" & idFactura & "]', '[" & fFactura & "]', '[" & nFactura & "]', '[ccFacturas_" & Year(fFactura) & "_Iva]', '')"
                                        
                                End Select
                            Next
                        End If
                    Next
                End If
            Next
            
            'Pedido
            If UCase(empresa) = UCase("Fac_Hitrs") Then
                If Not IsNumeric(nPedido) Then
                    If InStr(nPedido, " ") Then nPedido = Split(nPedido, " ")(0)
                End If
                
                ExecutaComandaSql "update [WEB]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] set reservat='XML:" & nPedido & "' where idFactura='" & idFactura & "'"

                If IsNumeric(nPedido) Then
                    Set rsPedido = Db.OpenResultset("select * from [web]." & empresa & ".dbo.DBA_Pedidos where id=" & nPedido) 'Cliente y fecha de pedido
                    If Not rsPedido.EOF Then
                        cliente = rsPedido("idCliente")
                        fPedido = rsPedido("fechaPedido")
                    Else
                        cliente = "2833"  'Hit Systems
                        'fPedido = Now()
                        fPedido = fFactura
                    End If
                Else
                    cliente = "2833" 'Hit Systems
                    'fPedido = Now()
                    fPedido = fFactura
                End If
                
                'Buscar cliente si es correctiva!!!
                If nFacturaCorrectiva <> "" Then
                    Set rsFactura = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.[" & tablaFacturaProforma(fFacturaCorrectiva) & "]  where numFactura='" & nFacturaCorrectiva & "'")
                    If Not rsFactura.EOF Then
                        Set rsCliente = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.[" & NomTaulaFacturaData(fFacturaCorrectiva) & "] where referencia like '%XML:" & rsFactura("Id") & "%'")
                        If Not rsCliente.EOF Then
                            cliente = rsCliente("Client")
                        End If
                    End If
                End If
        
                ExecutaComandaSql "Insert Into [web]." & empresa & ".dbo.FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values ('SincroMURANOFacturaRebuda', 0, '[" & idFactura & "]', '[" & fFactura & "]', '[" & nFactura & "]', '[ccFacturas_" & Year(fFactura) & "_Iva]', '')"
        
                impTotal = 0
                repartirImp = 0
                'Cogemos solo los 10 primeros dígitos del nombre del artículo porque a veces ponen nombres diferentes para el mismo código de producto y lo que nos interesa es el código que está al principio del nombre
                Set rsFactura = Db.OpenResultset("select distinct  f.idFactura, f.Data, f.Client, f.producte, left(producteNom, 10) producteNom, f.acabat, f.Preu, f.import, f.Desconte, f.tipusiva, f.iva, f.rec, f.referencia, f.servit, f.tornat, isnull(t.nLot, '') nLot from [WEB]." & empresa & ".dbo.[" & tablaFacturaProformaData(fFactura) & "] f left join [WEB]." & empresa & ".dbo.[INTERPRETA_XML_TMP] t on t.idfactura=f.idfactura and t.idproducto=f.producte where f.idFactura = '" & idFactura & "'")
                While Not rsFactura.EOF
                    Set rs = Db.OpenResultset("select codiArticle from [web]." & empresa & ".dbo.articlesPropietats where variable='MatPri'  and valor='" & rsFactura("producte") & "'")
                    If InStr(rsFactura("ProducteNom"), "4367") Or InStr(rsFactura("ProducteNom"), "21526") Or InStr(rsFactura("ProducteNom"), "21492") Or InStr(rsFactura("ProducteNom"), "21491") Or InStr(rsFactura("ProducteNom"), "21486") Or InStr(rsFactura("ProducteNom"), "21568") Or InStr(rsFactura("ProducteNom"), "21494") Or InStr(rsFactura("ProducteNom"), "22061") Or InStr(rsFactura("ProducteNom"), "1825") Then
                        preuArticle = rsFactura("Preu")
                    Else
                        preuArticle = rsFactura("Preu") * 1.35
                    End If
                    
                    If Not rs.EOF Then
                        codiArticle = rs("CodiArticle")
                        If preuArticle > 0 Then
                            ExecutaComandaSql "update [web]." & empresa & ".dbo.Articles set preu=" & (preuArticle) * (1 + (rsFactura("iva") / 100)) & ", preuMajor=" & preuArticle & " where codi=" & codiArticle
                        End If
                    Else
                        nomArt = rsFactura("ProducteNom")
                        Set rsNomArt = Db.OpenResultset("select top 1 producteNom from [WEB]." & empresa & ".dbo.[" & tablaFacturaProformaData(fFactura) & "] f where f.idFactura = '" & idFactura & "' and f.producte ='" & rsFactura("producte") & "'")
                        If Not rsNomArt.EOF Then nomArt = rsNomArt("producteNom")
                        
                        Set rs = Db.OpenResultset("select max(c) + 1 codi from (select max(codi) c from [web]." & empresa & ".dbo.articles union select max(codi) c from [web]." & empresa & ".dbo.articles_zombis) k")
                        If Not rs.EOF Then codiArticle = rs("codi")
        
                        ExecutaComandaSql "insert into [web]." & empresa & ".dbo.Articles (codi, nom, codiGenetic, esSumable, tipoIva, preu, preuMajor, nodescontesespecials) values(" & codiArticle & ",'" & nomArt & "'," & codiArticle & ", 1, " & rsFactura("TipusIva") & ", " & (preuArticle) * (1 + (rsFactura("iva") / 100)) & ", " & preuArticle & ", 0)"
                        ExecutaComandaSql "insert into [web]." & empresa & ".dbo.ArticlesPropietats (codiArticle, variable, valor) values (" & codiArticle & ", 'MatPri', '" & rsFactura("producte") & "')"
                    End If
                     
                    If InStr(rsFactura("ProducteNom"), "14715") Then
                        repartirImp = Round(rsFactura("Import") * 1.35, 2)
                    ElseIf InStr(rsFactura("ProducteNom"), "4367") Or InStr(rsFactura("ProducteNom"), "21526") Or InStr(rsFactura("ProducteNom"), "21492") Or InStr(rsFactura("ProducteNom"), "21491") Or InStr(rsFactura("ProducteNom"), "21486") Or InStr(rsFactura("ProducteNom"), "21568") Or InStr(rsFactura("ProducteNom"), "21494") Or InStr(rsFactura("ProducteNom"), "22061") Or InStr(rsFactura("ProducteNom"), "1825") Then
                        cliente = "2337" '2337    SILEMA BCN S.L.
                        ExecutaComandaSql "insert into [web]." & empresa & ".dbo.[Servit-" & Format(fPedido, "yy-mm-dd") & "] (Client, CodiArticle, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, TipusComanda, Comentari) values ('" & cliente & "', '" & codiArticle & "', 'Inicial', 'Inicial', " & rsFactura("servit") & ", 0, " & rsFactura("servit") & ", 3, '[XML:" & idFactura & "][Lote:" & rsFactura("nLot") & "]')"
                        impTotal = impTotal + Round(rsFactura("Import"), 2)
                        llistaArt = llistaArt & "," & codiArticle
                    Else
                        ExecutaComandaSql "insert into [web]." & empresa & ".dbo.[Servit-" & Format(fPedido, "yy-mm-dd") & "] (Client, CodiArticle, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, TipusComanda, Comentari) values ('" & cliente & "', '" & codiArticle & "', 'Inicial', 'Inicial', " & rsFactura("servit") & ", 0, " & rsFactura("servit") & ", 3, '[XML:" & idFactura & "][Lote:" & rsFactura("nLot") & "]')"
                        impTotal = impTotal + Round(rsFactura("Import") * 1.35, 2)
                        If InStr(1, llistaArt, "," & codiArticle) = 0 Then
                            llistaArt = llistaArt & "," & codiArticle
                        End If
                    End If
                    
                    rsFactura.MoveNext
                Wend
                
                'Repartir el importe del producto "21410 SERVEI PRE-POSTVENTA"
                If repartirImp > 0 And impTotal > 0 Then
                    pctReparto = Round(repartirImp / impTotal, 2)
                    llistaArtArr = Split(llistaArt, ",")
                    For a = 0 To UBound(llistaArtArr)
                        If llistaArtArr(a) <> "" Then ExecutaComandaSql "update [web]." & empresa & ".dbo.Articles set preu=preu + preu*" & pctReparto & ", preuMajor=preuMajor + preuMajor*" & pctReparto & " where codi=" & llistaArtArr(a)
                    Next
                End If
                
                If cliente <> "2833" And cliente <> "2850" Then 'No emite factura si es hit
                    'Pedimos la factura
                    sql = "Insert Into [web]." & empresa & ".dbo.FeinesAFer (Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5) "
                    sql = sql & "Values ('FesFactures', 0, 'Live Preus Actuals', "
                    sql = sql & "'[" & Right("0" & Day(fPedido), 2) & "-" & Right("0" & Month(fPedido), 2) & "-" & Year(fPedido) & " Al "
                    sql = sql & Right("0" & Day(fPedido), 2) & "-" & Right("0" & Month(fPedido), 2) & "-" & Year(fPedido) & "]', '[" & cliente & "]', "
                    sql = sql & "'[" & Right("0" & Day(Now()), 2) & "-" & Right("0" & Month(Now()), 2) & "-" & Year(Now()) & "]', "
                    sql = sql & "'[" & Right("0" & Day(Now()), 2) & "-" & Right("0" & Month(Now()), 2) & "-" & Year(Now()) & "]')"
                    ExecutaComandaSql sql
                End If
            End If
            
            sf_enviarMail "Secrehit@gmail.com", emailDe, "Factura XML interpretada " & nFactura, "Factura XML interpretada " & nFactura, "", ""
            'sf_enviarMail "Secrehit@gmail.com", "admin@hitsystems.es", "Factura XML interpretada " & nFactura, emailFactura_XML, "", ""
            'sf_enviarMail "Secrehit@gmail.com", "jordi@hit.cat", "Factura XML interpretada " & nFactura, emailFactura_XML, "", ""
            sf_enviarMail "Secrehit@gmail.com", "ana@solucionesit365.com", "Factura XML interpretada " & nFactura, emailFactura_XML, "", ""
        End If
    Else
        sf_enviarMail "Secrehit@gmail.com", emailDe, "LA FACTURA " & nFactura & " YA ESTABA IMPORTADA ", "", "", ""
    End If
        
    Exit Sub
    
ERR_XML:
    'FALTA BORRAR EL ALBARAN, PEDIDO, RECEPCION??????????????
    ExecutaComandaSql "Delete from [web]." & empresa & ".dbo." & tablaFacturaProforma(fFactura) & " where idFactura = '" & idFactura & "'"
    ExecutaComandaSql "Delete from [web]." & empresa & ".dbo." & tablaFacturaProformaData(fFactura) & " where idFactura = '" & idFactura & "'"
    
    sf_enviarMail "Secrehit@gmail.com", "ana@solucionesit365.com", "ERROR InterpretaFacturaXml " & nFactura, err.Description & "<br>" & nombreFichero & "<br>Fecha Pedido:[" & fPedido & "] Id recibida  [" & idFactura & "]  Cliente [" & cliente & "]", "", ""
    Exit Sub
    
ERR_XML_EMP:

    sf_enviarMail "Secrehit@gmail.com", emailDe, "NO ET CONEC", "", "", ""
End Sub

Sub InterpretaFacturaXml_Seller(NodeSellerParty As IXMLDOMNode, idFactura As String, fFactura As Date, empresa As String)
    Dim xNode As IXMLDOMNode, tNode As IXMLDOMNode, aNode As IXMLDOMNode
    Dim empNif As String, empNom As String, empAdresa As String, empCodi As String, tipoCobro As String
    Dim empCp As String, empCiutat As String, empTel As String, empFax As String, empEMail As String
    Dim rs As rdoResultset

    emailFactura_XML = emailFactura_XML & "<BR><B>EMISSOR</B><BR>"

    For Each xNode In NodeSellerParty.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "TaxIdentification"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "TaxIdentificationNumber"
                            InformaMiss "EMP NIF " & tNode.FirstChild.nodeValue & "<BR>"
                            empNif = tNode.FirstChild.nodeValue
                            emailFactura_XML = emailFactura_XML & "NIF: " & empNif & "<BR>"
                            'buscar "EmpresaCodi" en ccproveedores por nif
                            Set rs = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.ccproveedores where nif='" & empNif & "'")
                            If Not rs.EOF Then
                                empCodi = rs("id")
                                tipoCobro = rs("TipoCobro")
                            End If
                    End Select
                Next
            Case "LegalEntity"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "CorporateName"
                            InformaMiss "EMP NOM " & tNode.FirstChild.nodeValue & "<BR>"
                            empNom = tNode.FirstChild.nodeValue
                            emailFactura_XML = emailFactura_XML & "EMP NOM: " & empNom & "<BR>"
                        Case "AddressInSpain"
                            For Each aNode In tNode.childNodes
                                Select Case aNode.nodeName
                                    Case "Address"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP ADRESSA " & aNode.FirstChild.nodeValue & "<BR>"
                                            empAdresa = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "EMP ADRESA: " & empAdresa & "<BR>"
                                        End If
                                    Case "PostCode"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CP " & aNode.FirstChild.nodeValue & "<BR>"
                                            empCp = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "EMP CP: " & empCp & "<BR>"
                                        End If
                                    Case "Town"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            empCiutat = aNode.FirstChild.nodeValue
                                        End If
                                    Case "Province"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            empCiutat = empCiutat & " " & aNode.FirstChild.nodeValue
                                        End If
                                    Case "CountryCode"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            empCiutat = empCiutat & " " & aNode.FirstChild.nodeValue
                                        End If
                                End Select
                            Next
                            emailFactura_XML = emailFactura_XML & "EMP Ciutat: " & empCiutat & "<BR>"
                        Case "ContactDetails"
                            For Each aNode In tNode.childNodes
                                Select Case aNode.nodeName
                                    Case "Telephone"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP TEL " & aNode.FirstChild.nodeValue & "<BR>"
                                            empTel = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "EMP TEL: " & empTel & "<BR>"
                                        End If
                                    Case "TeleFax"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP FAX " & aNode.FirstChild.nodeValue & "<BR>"
                                            empFax = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "EMP FAX: " & empFax & "<BR>"
                                        End If
                                    Case "ElectronicMail"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP EMAIL " & aNode.FirstChild.nodeValue
                                            empEMail = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "EMP EMAIL: " & empEMail & "<BR>"
                                        End If
                                End Select
                            Next
                    End Select
                Next
        End Select
    Next

    ExecutaComandaSql "update [web]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] set EmpresaCodi='" & empCodi & "', EmpNif='" & empNif & "', EmpNom='" & empNom & "', EmpAdresa='" & empAdresa & "', EmpCp='" & empCp & "', EmpCiutat='" & empCiutat & "', EmpTel='" & empTel & "', EmpFax='" & empFax & "', EmpeMail='" & empEMail & "', FormaPagament='" & tipoCobro & "' where idFactura='" & idFactura & "'"
End Sub
Sub InterpretaFacturaXml_Proveedor(NodeSellerParty As IXMLDOMNode, ByRef idProveedor, empresa As String)
    Dim xNode As IXMLDOMNode, tNode As IXMLDOMNode, aNode As IXMLDOMNode
    Dim rs As rdoResultset, empNif As String

    For Each xNode In NodeSellerParty.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "TaxIdentification"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "TaxIdentificationNumber"
                            empNif = tNode.FirstChild.nodeValue
                            'buscar "EmpresaCodi" en ccproveedores por nif
                            Set rs = Db.OpenResultset("select * from [WEB]." & empresa & ".dbo.ccproveedores where nif='" & empNif & "'")
                            If Not rs.EOF Then
                                idProveedor = rs("id")
                            End If
                    End Select
                Next
        End Select
    Next
End Sub

Sub InterpretaFacturaXml_Empresa(NodeBuyerParty As IXMLDOMNode, ByRef empresa As String)
    Dim xNode As IXMLDOMNode, tNode As IXMLDOMNode, aNode As IXMLDOMNode
    Dim rs As rdoResultset, rsEmp As rdoResultset, empNif As String

    For Each xNode In NodeBuyerParty.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "TaxIdentification"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "TaxIdentificationNumber"
                            empNif = tNode.FirstChild.nodeValue
                            'buscar empresa, que recibe la factura de compras, por nif
                            Set rs = Db.OpenResultset("select * from sys.databases where name like 'Fac_%' and name not like '%bak%'")
                            While Not rs.EOF
                                Set rsEmp = Db.OpenResultset("select * from [WEB]." & rs("name") & ".dbo.constantsEmpresa where valor = '" & Trim(empNif) & "'")
                                If Not rsEmp.EOF Then
                                    empresa = rs("name")
                                    Exit Sub
                                End If
                                rs.MoveNext
                            Wend
                    End Select
                Next
        End Select
    Next
End Sub

Sub InterpretaFacturaXml_Buyer(NodeSellerParty As IXMLDOMNode, idFactura As String, fFactura As Date, empresa As String)
    Dim xNode As IXMLDOMNode, tNode As IXMLDOMNode, aNode As IXMLDOMNode
    Dim CliNif As String, CliNom As String, CliAdresa As String, Clicodi As String
    Dim CliCp As String, CliCiutat As String, cliTel As String, cliFax As String, CliEmail As String
    Dim rs As rdoResultset

    emailFactura_XML = emailFactura_XML & "<BR><B>RECPETOR</B><BR>"
    

    For Each xNode In NodeSellerParty.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "TaxIdentification"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "TaxIdentificationNumber"
                            InformaMiss "CLI NIF " & tNode.FirstChild.nodeValue & "<BR>"
                            CliNif = tNode.FirstChild.nodeValue
                            emailFactura_XML = emailFactura_XML & "NIF: " & CliNif & "<BR>"
                            'buscar "ClientCodi" en constantsempresa por nif
                            Clicodi = "0"
                            Set rs = Db.OpenResultset("select * from [web]." & empresa & ".dbo.constantsEmpresa where valor='" & CliNif & "' order by camp desc")
                            If Not rs.EOF Then
                                If InStr(rs("Camp"), "_") Then Clicodi = Split(rs("camp"), "_")(0)
                            End If
                    End Select
                Next
            Case "LegalEntity"
                For Each tNode In xNode.childNodes
                    Select Case tNode.nodeName
                        Case "CorporateName"
                            InformaMiss "EMP NOM " & tNode.FirstChild.nodeValue & "<BR>"
                            CliNom = tNode.FirstChild.nodeValue
                            emailFactura_XML = emailFactura_XML & "CLI NOM: " & CliNom & "<BR>"
                        Case "AddressInSpain"
                            For Each aNode In tNode.childNodes
                                Select Case aNode.nodeName
                                    Case "Address"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP ADRESSA " & aNode.FirstChild.nodeValue & "<BR>"
                                            CliAdresa = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "CLI ADRESA: " & CliAdresa & "<BR>"
                                        End If
                                    Case "PostCode"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CP " & aNode.FirstChild.nodeValue & "<BR>"
                                            CliCp = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "CLI CP: " & CliCp & "<BR>"
                                        End If
                                    Case "Town"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            CliCiutat = aNode.FirstChild.nodeValue
                                        End If
                                    Case "Province"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            CliCiutat = CliCiutat & " " & aNode.FirstChild.nodeValue
                                        End If
                                    Case "CountryCode"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP CIUTAT " & aNode.FirstChild.nodeValue
                                            CliCiutat = CliCiutat & " " & aNode.FirstChild.nodeValue
                                        End If
                                End Select
                            Next
                            emailFactura_XML = emailFactura_XML & "CLI Ciutat: " & CliCiutat & "<BR>"
                        Case "ContactDetails"
                            For Each aNode In tNode.childNodes
                                Select Case aNode.nodeName
                                    Case "Telephone"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP TEL " & aNode.FirstChild.nodeValue & "<BR>"
                                            cliTel = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "CLI TEL: " & cliTel & "<BR>"
                                        End If
                                    Case "TeleFax"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP FAX " & aNode.FirstChild.nodeValue & "<BR>"
                                            cliFax = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "CLI FAX: " & cliFax & "<BR>"
                                        End If
                                    Case "ElectronicMail"
                                        If aNode.hasChildNodes Then
                                            InformaMiss "EMP EMAIL " & aNode.FirstChild.nodeValue
                                            CliEmail = aNode.FirstChild.nodeValue
                                            emailFactura_XML = emailFactura_XML & "CLI EMAIL: " & CliEmail & "<BR>"
                                        End If
                                End Select
                            Next
                    End Select
                Next
        End Select
    Next

    ExecutaComandaSql "update [web]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] set ClientCodi=" & Clicodi & ", ClientCodiFac=" & Clicodi & ", ClientNif='" & CliNif & "', ClientNom='" & CliNom & "', ClientAdresa='" & CliAdresa & "', ClientCp='" & CliCp & "', ClientCiutat='" & CliCiutat & "', Tel='" & cliTel & "', Fax='" & cliFax & "', eMail='" & CliEmail & "' where idFactura='" & idFactura & "'"
End Sub

Sub InterpretaFacturaXml_Taxes(iNode As IXMLDOMNode, idFactura As String, fFactura As Date, empresa As String)
    Dim xNode As IXMLDOMNode, tNode As IXMLDOMNode
    Dim pIVA As Double, bIVA As Double, cIVA As Double
    Dim pIVA1 As Double, bIVA1 As Double, cIVA1 As Double, pIVA2 As Double, bIVA2 As Double, cIVA2 As Double, pIVA3 As Double, bIVA3 As Double, cIVA3 As Double

    pIVA1 = 0: bIVA1 = 0: cIVA1 = 0
    pIVA2 = 0: bIVA2 = 0: cIVA2 = 0
    pIVA3 = 0: bIVA3 = 0: cIVA3 = 0
    
    emailFactura_XML = emailFactura_XML & "<BR><B>TAXES</B><BR>"

    For Each xNode In iNode.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "Tax"
                For Each tNode In xNode.childNodes
                    If tNode.nodeName = "TaxRate" Then
                        InformaMiss "IVA " & tNode.FirstChild.nodeValue
                        pIVA = CDbl(tNode.FirstChild.nodeValue)
                    ElseIf tNode.nodeName = "TaxableBase" Then
                        InformaMiss "BASE " & tNode.FirstChild.FirstChild.nodeValue
                        bIVA = CDbl(tNode.FirstChild.FirstChild.nodeValue)
                    ElseIf tNode.nodeName = "TaxAmount" Then
                        InformaMiss "CUOTA " & tNode.FirstChild.FirstChild.nodeValue
                        cIVA = CDbl(tNode.FirstChild.FirstChild.nodeValue)
                    End If
                Next
                If pIVA = 4 Then pIVA1 = pIVA: bIVA1 = bIVA: cIVA1 = cIVA
                If pIVA = 10 Then pIVA2 = pIVA: bIVA2 = bIVA: cIVA2 = cIVA
                If pIVA = 21 Then pIVA3 = pIVA:  bIVA3 = bIVA: cIVA3 = cIVA

                emailFactura_XML = emailFactura_XML & "IVA: " & pIVA & "% BASE: " & bIVA & " CUOTA: " & cIVA & "<BR>"
        End Select
    Next

    ExecutaComandaSql "update [web]." & empresa & ".dbo.[" & tablaFacturaProforma(fFactura) & "] set BaseIva1=" & bIVA1 & ", Iva1=" & cIVA1 & ", BaseIva2=" & bIVA2 & ", Iva2=" & cIVA2 & ", BaseIva3=" & bIVA3 & ", Iva3=" & cIVA3 & " where idFactura='" & idFactura & "'"

End Sub




Sub InterpretaFacturaXml_Linea(iNode As IXMLDOMNode, idFactura As String, nFactura As String, fFactura As Date, idProveedor As String, ByRef nPedido As String, empresa As String)
    Dim xNode As IXMLDOMNode, dNode As IXMLDOMNode
    Dim fecha As Date, nAlb As String, prodNom As String, prodCodi As String, prodId As String, contrapartida As String
    Dim servit As Double, Preu As Double, import As Double, tipusIva As Integer, iva As Integer, dte As Double
    Dim sql As String
    Dim rs As rdoResultset
    Dim rsProd As rdoResultset, rsIdPedido As rdoResultset
    Dim nSerie As String, idAlmacen As String, idPedido As String
    Dim rsAgrupar As rdoResultset, agrupar As Boolean
    
    For Each xNode In iNode.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "FileReference"
                nSerie = ""
                If xNode.hasChildNodes Then
                    nSerie = xNode.FirstChild.nodeValue
                End If
                prodNom = prodNom & " " & xNode.FirstChild.nodeValue
                emailFactura_XML = emailFactura_XML & "PROD ADD INF: " & prodNom & "<BR>"
            Case "SequenceNumber"
                InformaMiss "SequenceNumber: " & xNode.FirstChild.nodeValue, True
            Case "DeliveryNotesReferences"
                For Each dNode In xNode.FirstChild.childNodes
                    If dNode.nodeName = "DeliveryNoteNumber" Then nAlb = dNode.FirstChild.nodeValue
                    If dNode.nodeName = "DeliveryNoteDate" Then fecha = dNode.FirstChild.nodeValue
                Next
                emailFactura_XML = emailFactura_XML & "<BR><B>N ALB:</B> " & nAlb & " FECHA: " & fecha & "<BR>"
            Case "ItemDescription"
                prodNom = "Desconocido"
                If xNode.hasChildNodes Then
                    prodNom = xNode.FirstChild.nodeValue
                    If InStr(prodNom, "COMANDA #") Then
                        nPedido = Split(prodNom, "COMANDA #")(1)
                    End If
                    If InStr(prodNom, "COMANDA#") Then
                        nPedido = Split(prodNom, "COMANDA#")(1)
                    End If

                    prodNom = Left(prodNom, 255)
                End If
                emailFactura_XML = emailFactura_XML & "PROD: " & prodNom & "<BR>"
            Case "AdditionalLineItemInformation" 'n serie
                prodNom = prodNom & " " & xNode.FirstChild.nodeValue
                emailFactura_XML = emailFactura_XML & "PROD ADD INF: " & prodNom & "<BR>"
            Case "ArticleCode"
                prodCodi = xNode.FirstChild.nodeValue
                emailFactura_XML = emailFactura_XML & "CODIGO PROD: " & prodCodi & "<BR>"
            Case "Quantity"
                servit = CDbl(xNode.FirstChild.nodeValue)
                emailFactura_XML = emailFactura_XML & "SERVIT: " & prodNom & "<BR>"
            Case "UnitPriceWithoutTax"
                If xNode.hasChildNodes Then
                    Preu = CDbl(xNode.FirstChild.nodeValue)
                End If
                emailFactura_XML = emailFactura_XML & "PREU: " & Preu & "<BR>"
            Case "TotalCost"
                If xNode.hasChildNodes Then
                    import = CDbl(xNode.FirstChild.nodeValue)
                End If
                emailFactura_XML = emailFactura_XML & "IMPORT: " & import & "<BR>"
            Case "GrossAmount"    'Hay descuentos
                If xNode.hasChildNodes Then
                    import = CDbl(xNode.FirstChild.nodeValue)
                End If
                emailFactura_XML = emailFactura_XML & "IMPORT DTE: " & import & "<BR>"
            Case "DiscountsAndRebates"
                For Each dNode In xNode.FirstChild.childNodes
                    If dNode.nodeName = "DiscountRate" Then dte = CDbl(dNode.FirstChild.nodeValue)
                Next
            Case "TaxesOutputs"
                iva = 21
                tipusIva = 3
                For Each dNode In xNode.FirstChild.childNodes
                    If dNode.nodeName = "TaxRate" Then iva = CInt(dNode.FirstChild.nodeValue)
                Next

                Set rs = Db.OpenResultset("select tipus from [web]." & empresa & ".dbo." & DonamTaulaTipusIva(Now()) & " where iva = " & iva)
                If Not rs.EOF Then tipusIva = rs("tipus")
        End Select
    Next
        
    prodNom = Left(Replace(Replace(prodNom, """", " "), "'", "´"), 255)
    
    agrupar = False
    Set rsAgrupar = Db.OpenResultset("select * from [web]." & empresa & ".dbo.ccProveedoresExtes where nom = 'FacturasAgrupadas' and valor='on' and id='" & idProveedor & "'")
    If Not rsAgrupar.EOF Then agrupar = True


    'PARA FACTURA DETALLADA SIN AGRUPAR (HITRS, SOLUCIONES) ----------------------------------------------------------------------------------------------------------------------
    If Not agrupar Then
    
        prodId = prodCodi
        
        Set rs = Db.OpenResultset("select * from [web]." & empresa & ".dbo.ccNombreValor where nombre='Refinterna' and valor='" & prodCodi & "'")
        If Not rs.EOF Then
            prodId = rs("id")
            Set rs = Db.OpenResultset("select * from [web]." & empresa & ".dbo.ccMateriasPrimas where id='" & rs("id") & "'")
            If Not rs.EOF Then
                prodId = rs("id")
            Else
                ExecutaComandaSql "insert into [web]." & empresa & ".dbo.ccMateriasPrimas (id, nombre, activo, proveedor, precio, precioFormato, Iva) values ('" & prodId & "', '" & prodNom & "', 1, '" & idProveedor & "', " & Preu & ", " & Preu & ", 3)"
            End If
        Else   'Crear el producto
            Set rs = Db.OpenResultset("select newid() Id")
            If Not rs.EOF Then prodId = rs("id")
            
            ExecutaComandaSql "insert into [web]." & empresa & ".dbo.ccMateriasPrimas (id, nombre, activo, proveedor, precio, precioFormato, Iva) values ('" & prodId & "', '" & prodNom & "', 1, '" & idProveedor & "', " & Preu & ", " & Preu & ", 3)"
            ExecutaComandaSql "insert into [web]." & empresa & ".dbo.ccNombreValor (id, nombre, valor) values ('" & prodId & "', 'Refinterna', '" & prodCodi & "')"
        End If
    
        contrapartida = ""
        
        Set rs = Db.OpenResultset("select valor from [web]." & empresa & ".dbo.ccNombreValor where id = '" & prodId & "' and nombre='Contrapartida'")
        If Not rs.EOF Then contrapartida = rs("valor")
        
        If contrapartida = "" Then
            Set rs = Db.OpenResultset("select valor from [web]." & empresa & ".dbo.ccproveedoresextes where id='" & idProveedor & "' and nom like '%Contrapartida%' order by valor")
            If Not rs.EOF Then
                If InStr(rs("valor"), "|") Then contrapartida = Split(rs("valor"), "|")(1)
            End If
        End If
        
        'If Left(Trim(prodNom), 4) = "4367" Then
        '    contrapartida = "607000000"
        'End If
        
        If import <> 0 Then
            If fecha = "0:00:00" Then fecha = fFactura
            
            'PEDIDO
            idAlmacen = "" 'DE MOMENTO NO HAY ALAMACÉN
            Set rsIdPedido = Db.OpenResultset("select newid() IdPedido from [web]." & empresa & ".dbo.[ccPedidos]")
            idPedido = rsIdPedido("IdPedido")
            
            'ExecutaComandaSql "insert into " & NomTaulaPedidosCap() & " (numPedido, idPedido, proveedor, almacen, fecha, recepcion) values (" & nPedido & ", '" & idPedido & "', '" & idProveedor & "', '', getdate(), getdate())"
                        
            sql = "insert into [web]." & empresa & ".dbo.[ccPedidos] (Id, MateriaPrima, Proveedor, Almacen, cantidad, fecha, recepcion, precio, activo) "
            sql = sql & "values ('" & idPedido & "', '" & prodId & "', '" & idProveedor & "', '" & idAlmacen & "', " & servit & ", convert(datetime, '" & fecha & "', 103), convert(datetime, '" & fecha & "', 103), " & Preu & ", 1)"
            ExecutaComandaSql sql
            
            sql = "insert into [web]." & empresa & ".dbo.[ccPedidosExtes] (IdPedido, Variable, Valor) values ('" & idPedido & "', 'FACTURA', '" & nFactura & "')"
            ExecutaComandaSql sql
            
            sql = "insert into [web]." & empresa & ".dbo.[ccPedidosExtes] (IdPedido, Variable, Valor) values ('" & idPedido & "', 'ID_FACTURA', '" & idFactura & "')"
            ExecutaComandaSql sql
            
            'RECEPCIÓN
            sql = "insert into [web]." & empresa & ".dbo.[ccRecepcion] (Id , proveedor, MatPrima, albaran, pedido, temperatura, caract, envas, Usuario, fecha, aceptado, Lote, facturado, caducidad) "
            sql = sql & "values (newid(), '" & idProveedor & "', '" & prodId & "', '" & nAlb & "', '" & idPedido & "', 0, 0, 0, 'SINCRO', convert(datetime, '" & fecha & "', 103), 0, '" & nSerie & "', 1, convert(datetime, '" & DateAdd("m", 1, fecha) & "', 103))"
            ExecutaComandaSql sql

            'FACTURA
            sql = "insert into [web]." & empresa & ".dbo.[" & tablaFacturaProformaData(fFactura) & "] (IdFactura, Data, Producte, ProducteNom, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) "
            sql = sql & "values ('" & idFactura & "', convert(datetime, '" & fecha & "', 103), '" & prodId & "', '" & prodNom & "', " & Preu & ", " & import & ", " & dte & ", " & tipusIva & ", " & iva & ", 0, '" & contrapartida & "', " & servit & ", 0)"
            ExecutaComandaSql sql
            
            ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.[INTERPRETA_XML_TMP] (idFactura, idProducto, nLot) values ('" & idFactura & "', '" & prodId & "', '" & nSerie & "')"
            
        End If
        
        'SI HAY NÚMERO DE SERIE GUARDAMOS EL PRODUCTO EN RECURSOS
        If nSerie <> "" Then
            Dim rsRecId As rdoResultset, idRecurso As String
            Dim rsPedido As rdoResultset, rsExisteix As rdoResultset
            Set rsRecId = Db.OpenResultset("select newid() Id")
            idRecurso = rsRecId("Id")
            
            ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.recursos (id, nombre, tipo) values ('" & idRecurso & "', '" & Left(prodNom, 255) & "', 'HARDWARE')"
            ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.recursosExtes (id, variable, valor) values ('" & idRecurso & "', 'DESCRIPCION', '" & Left(prodNom, 255) & "')"
            ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.recursosExtes (id, variable, valor) values ('" & idRecurso & "', 'NSERIE', '" & nSerie & "')"
            ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.recursosExtes (id, variable, valor) values ('" & idRecurso & "', 'NPEDIDO', '" & nPedido & "')"
            
            Set rsExisteix = Db.OpenResultset("select * from [WEB]." & empresa & ".sys.tables where name ='dba_pedidos'")
            If Not rsExisteix.EOF Then
                Set rsPedido = Db.OpenResultset("select c.nom, c.codi from [WEB]." & empresa & ".dbo.dba_pedidos p left join [WEB]." & empresa & ".dbo.clients c on p.idCliente = c.codi where p.id=" & nPedido)
                If Not rsPedido.EOF Then
                    ExecutaComandaSql "insert into [WEB]." & empresa & ".dbo.recursosExtes (id, variable, valor) values ('" & idRecurso & "', 'CLIENTE', '" & rsPedido("codi") & "')"
                End If
            End If
        End If
        
    Else 'PARA FACTURAS AGRUPADAS (IME, SILEMA, ...) -------------------------------------------------------------------------------------------------------------------
        'Miramos si el producto de venta (que es lo que viene en la factura) tiene una materia prima asociada en compras
        ExecutaComandaSql "Delete from [web]." & empresa & ".dbo.articlespropietats where variable='MatPri' and valor=''"
        
        Set rs = Db.OpenResultset("select * from [web]." & empresa & ".dbo.ccMateriasPrimas where id in (select isnull(valor, '-') pId from [web]." & empresa & ".dbo.articlespropietats where codiarticle=" & prodCodi & " and variable='MatPri')")
        If Not rs.EOF Then
            prodNom = rs("nombre")
            prodId = rs("Id")
        Else
            Set rs = Db.OpenResultset("select newid() Id")
            If Not rs.EOF Then prodId = rs("id")
        End If
        
        contrapartida = "60000000"
        If nAlb = "ALB_0000IBEE" Then
            contrapartida = "63101000"
        End If
        'Set rs = Db.OpenResultset("select valor from [web]." & Empresa & ".dbo.ccproveedoresextes where id='" & idProveedor & "' and nom like '%Contrapartida%'")
        'If Not rs.EOF Then
        '    If InStr(rs("valor"), "|") Then contrapartida = Split(rs("valor"), "|")(1)
        'End If
        
        If import <> 0 Then
            If fecha = "0:00:00" Then fecha = fFactura
            sql = "insert into [web]." & empresa & ".dbo.[ccFacturasData_XML] (IdFactura, Data, Producte, ProducteNom, Preu, Import, Desconte, TipusIva, Iva, Rec, Referencia, Servit, Tornat) "
            sql = sql & "values ('" & idFactura & "', convert(datetime, '" & fecha & "', 103), '" & prodId & "', '" & prodNom & "', " & Round(Preu * servit, 3) & ", " & Round(import * servit, 3) & ", " & dte & ", " & tipusIva & ", " & iva & ", 0, '" & contrapartida & "', " & servit & ", 0)"
            ExecutaComandaSql sql
        End If
    End If

End Sub


Sub InterpretaFacturaXml_Header(NodeInvoice As IXMLDOMNode, ByRef fFactura As Date, ByRef fVencimiento As Date, ByRef nFactura As String, ByRef nFacturaCorrectiva As String, ByRef fFacturaCorrectiva As Date)
    Dim xNode As IXMLDOMNode, xNodeC As IXMLDOMNode
    
    For Each xNode In NodeInvoice.childNodes
        InformaMiss "Nodo: " & xNode.nodeName, True
        Select Case xNode.nodeName
            Case "InvoiceHeader"
                InformaMiss "NUMERO DE FACTURA " & xNode.FirstChild.FirstChild.nodeValue 'InvoiceHeader-InvoiceNumber
                nFactura = Replace(xNode.FirstChild.FirstChild.nodeValue, "//", "/")
            Case "InvoiceIssueData"
                InformaMiss "FECHA FACTURA " & xNode.FirstChild.FirstChild.nodeValue 'InvoiceIssueData-IssueDate
                fFactura = xNode.FirstChild.FirstChild.nodeValue
            Case "PaymentDetails"
                InformaMiss "FECHA VENCIMIENTO " & xNode.FirstChild.FirstChild.FirstChild.nodeValue 'Installment-InstallmentDueDate
                fVencimiento = xNode.FirstChild.FirstChild.FirstChild.nodeValue
            Case "Corrective"
                For Each xNodeC In xNode.childNodes
                    Select Case xNodeC.nodeName
                        Case "InvoiceNumber"
                            nFacturaCorrectiva = xNodeC.FirstChild.nodeValue
                        Case "TaxPeriod"
                            fFacturaCorrectiva = xNodeC.FirstChild.FirstChild.nodeValue
                    End Select
                Next
        End Select
    Next
    
End Sub


Sub RevisaEmailDe(User As String, Password As String, Optional filtre As String = "")
    Dim UnDeBorrat As Boolean, NomAtt As String, H1 As Date, i, downHTTP As New HTTP, empresa As String, subj As String, t, iD, nomF, ruta1, ruta2, ruta3, sql, extF, mimeF, NomUsu, n, rs, URL, resposta, aResposta, emailDe As String, IdFoto As String, Borrar As Boolean, Text As String
    Dim fecha As Date, HTMLText As String, nombreFichero As String
    Dim nomFileTemp As String
    Dim objStream As Object
    
On Error GoTo nor
    empresa = ""
    
    InformaEmpresa "Hit Email " & User
    Set frmSplash.POP3 = New wodPop3Com
    frmSplash.POP3.HostName = "pop.gmail.com"
    frmSplash.POP3.Security = SecurityImplicit
    
    frmSplash.POP3.Port = 995
    frmSplash.POP3.Login = User
    frmSplash.POP3.Password = Password
    frmSplash.POP3.Connect
   
    Informa "Conectant " & frmSplash.POP3.Login
    H1 = Now
    While Not frmSplash.POP3.State = WODPOP3COMLib.StatesEnum.Connected
        DoEventsSleep
        If Now > DateAdd("s", 90, H1) Then
            Informa "Se acabó el tiempo"
            Exit Sub
        End If
    Wend
    Informa "Conectat " & frmSplash.POP3.Messages.Count & " Emails"
    
    For i = 0 To frmSplash.POP3.Messages.Count - 1
        Set frmSplash.Pop3Message = Nothing
        Borrar = False
        Informa "Llegint " & i & " De " & frmSplash.POP3.Messages.Count & " "
        frmSplash.POP3.Messages(i).Get
        H1 = Now
        DoEventsSleep
        While frmSplash.Pop3Message Is Nothing
            DoEventsSleep
            If Now > DateAdd("s", 90, H1) Then
                Exit Sub
            End If
        Wend
        
        Informa frmSplash.Pop3Message.Subject '& "(" & frmSplash.Pop3Message.Attachments.Count & ")"
        emailDe = frmSplash.Pop3Message.FromEmail
        empresa = SecreHitEmailEmpresa(emailDe)
        Text = frmSplash.Pop3Message.PlainText
        HTMLText = frmSplash.Pop3Message.HTMLText
        
        ExecutaComandaSql "Insert into hit.dbo.DebugEmailSecres (Secre, Empresa, Usuario, Asunto, Fecha) values ('" & User & "', '" & empresa & "', '" & emailDe & "', '" & Mid(frmSplash.Pop3Message.Subject, 1, 255) & "', getdate())"
        
        'COMPROBAR SI HAY XML ADJUNTO!!!
        Dim a As Integer
        For a = 0 To frmSplash.Pop3Message.Attachments.Count - 1
            If InStr(frmSplash.Pop3Message.Attachments(a).Name, ".") > 0 Then
                extF = Right(frmSplash.Pop3Message.Attachments(a).Name, Len(frmSplash.Pop3Message.Attachments(a).Name) - InStr(frmSplash.Pop3Message.Attachments(a).Name, "."))
                If UCase(extF) = "XML" Then
                    NomAtt = "AttEmail." & extF
                    MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                    frmSplash.Pop3Message.Attachments(a).Save "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                    nombreFichero = frmSplash.Pop3Message.Attachments(a).Name
                    Informa "Rebut File " & frmSplash.Pop3Message.Attachments(a).Name

                    Exit For
                End If
            End If
        Next
        
        'SI NOS ENVIAN UN XML INTENTAMOS INTERPRETARLO
        If UCase(extF) = "XML" Then
            t = "Factura"
            InterpretaFacturaXml "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt, empresa, emailDe, nombreFichero
            ExecutaComandaSql "Insert into hit.dbo.DebugEmailSecres (Secre, Empresa, Usuario, Asunto, Fecha) values ('" & User & "', '" & empresa & "', '" & emailDe & "', 'FICHERO XML: " & Right(nombreFichero, 200) & "', getdate())"
            Borrar = True
            MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
            GoTo NEXT_MSG
        End If
       
        If filtre = "" Then
            Borrar = True
            
            subj = frmSplash.Pop3Message.Subject
            If empresa = "" And emailDe = "email@hit.cat" Then
                If Left(subj, 18) = "INCIDENCIA CERBERO" Then
                    empresa = Mid(subj, InStr(subj, "[") + 1, InStr(subj, "]") - InStr(subj, "[") - 1)
                End If
            End If

            'NO SÉ QUIEN ME ENVIA EL EMAIL. NO ESTÁ DADO DE ALTA EN LA SECRE (Hit.dbo.Secretaria)
            If empresa = "" Then
                EnviaEmail emailDe, "NO ET CONEC"
            Else
                ExecutaComandaSql "use  " & empresa
                
                'NO HAY DOCUMENTO ADJUNTO --------------------------------------------------------------------------------------------------------------------
                'If frmSplash.Pop3Message.Attachments.Count = 0 Then
                                        
                    'INFORME DE INFORMES (HELP)
                    If UCase(subj) = UCase("help") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeHelp', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE LICENCIAS HIT
                    ElseIf UCase(subj) = "INFORME LICENCIAS" Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeLicencias', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE VENTAS
                    ElseIf UCase(Left(subj, 14)) = UCase("informe ventas") Or UCase(Left(subj, 14)) = UCase("informe_ventas") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeVentas', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE REPOSICIÓN AUTOMÁTICA
                    ElseIf UCase(subj) = "INFORME REPOSICION" Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeReposicion', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME SUPERVISORA
                    ElseIf UCase(Left(subj, 19)) = UCase("informe supervisora") Or UCase(Left(subj, 19)) = UCase("informe_supervisora") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeSupervisora', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME COORDINADORA
                    ElseIf UCase(Left(subj, 20)) = UCase("informe coordinadora") Or UCase(Left(subj, 20)) = UCase("informe_coordinadora") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeCoordinadora', 0, '" & Replace(Replace(subj, "_", " ") & "','", "—", "--") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'RESPUESTA AL INFORME DE COORDINADORA
                    ElseIf UCase(Left(subj, 24)) = UCase("Re: Informe coordinadora") Then
                     If UCase(emailDe) = UCase("estefani-estar@hotmail.com") Or UCase(emailDe) = UCase("mariasheyla@hotmail.es") Or UCase(emailDe) = UCase("montse_pg@hotmail.com") Or UCase(emailDe) = UCase("si_mo21@hotmail.com") Or UCase(emailDe) = UCase("vajuliet@hotmail.com") Or UCase(emailDe) = UCase("pao_pio77@hotmail.com") Then
                        nomFileTemp = "RESPUESTA_COORDINADORA_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhnnss") & "000.HTML"
                        'Create the stream
                        Set objStream = CreateObject("ADODB.Stream")
                        
                        'Initialize the stream
                        objStream.Open
                        
                        'Reset the position and indicate the charactor encoding
                        objStream.Position = 0
                        objStream.Charset = "UTF-8"

                        'Write to the steam
                        objStream.WriteText HTMLText
                        
                        'Save the stream to a file
                        objStream.SaveToFile "c:\" & nomFileTemp
                      End If
                        preparaCodigoDeAccionCuadrante HTMLText, emailDe, empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeCoordinadora', 0, '" & Replace(Replace(Replace(subj, "_", " ") & "','", "—", "--"), "Re: ", "") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME RESUMEN VENTAS
                    ElseIf UCase(Left(subj, 14)) = UCase("resumen ventas") Or UCase(Left(subj, 14)) = UCase("resumen_ventas") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeResumenVentas', 0, '" & Replace(Replace(subj, "_", " ") & "','", "—", "--") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME INSTALACIÓN TIENDA
                    ElseIf UCase(Left(subj, 11)) = UCase("instalacion") Or UCase(Left(subj, 11)) = UCase("instalación") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeInstalacion', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME TRASPASO
                    ElseIf UCase(Left(subj, 16)) = UCase("informe traspaso") Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeTraspaso', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE PEDIDO SEMANAL
                    ElseIf UCase(Left(subj, 14)) = UCase("pedido semanal") Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformePedidoSemanal', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE PRODUCTOS TOP
                    ElseIf UCase(Left(subj, 17)) = "INFORME PRODUCTOS" Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeProductos', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'LISTADO PRODUCTOS EN USO
                    ElseIf UCase(subj) = UCase("LISTADO PRODUCTOS EN USO") Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeUsoProductos', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'LISTADO PRODUCTOS / MATERIA PRIMA
                    ElseIf UCase(subj) = UCase("PRODUCTOS COMPRA VENTA") Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeProductosCompraVenta', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE PREVISIONES
                    'ElseIf InStr(UCase(subj), UCase("informe previsiones")) Then
                    '    ExecutaComandaSql "use  " & Empresa
                    '    ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'calculaInformePrevisiones', 0, '" & subj & "','" & emailDe & "', '" & Empresa & "', '', '', getdate())"
                                        
                    ElseIf InStr(UCase(subj), UCase("informe previsiones")) Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'calculaInformePrevisiones2', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'ACTUALIZA PREVISIONES
                    ElseIf InStr(UCase(subj), UCase("previsiones semana")) Then
                        ExecutaComandaSql "use  " & empresa
                        
                        nomFileTemp = "PREVISIONES_SEMANA_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhnnss") & "000.HTML"
                        'Create the stream
                        Set objStream = CreateObject("ADODB.Stream")
                        
                        'Initialize the stream
                        objStream.Open
                        
                        'Reset the position and indicate the charactor encoding
                        objStream.Position = 0
                        objStream.Charset = "UTF-8"

                        'Write to the steam
                        objStream.WriteText HTMLText
                        
                        'Save the stream to a file
                        objStream.SaveToFile "c:\" & nomFileTemp
    
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreActualizaPrevisiones', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', 'c:\" & nomFileTemp & "', '', getdate())"
                        'actualizaPrevisiones Empresa, emailDe, HTMLText
                        GoTo NEXT_MSG
                    'INFORME DE INCIDENCIAS
                    ElseIf UCase(subj) = UCase("informe incidencias") Then
                        ExecutaComandaSql "use " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeIncidencias', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME DE INCIDENCIAS SILEMA
                    'ElseIf InStr(UCase(subj), UCase("incidencias silema")) Then
                    '    ExecutaComandaSql "use " & Empresa
                    '    ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeIncidenciasSilema', 0, '" & subj & "','" & emailDe & "', '" & Empresa & "', '', '', getdate())"
                    '    GoTo NEXT_MSG
                    'INFORME MASAS
                    ElseIf Left(UCase(subj), 13) = UCase("informe masas") Then
                        ExecutaComandaSql "use " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeMasas', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME EBITDA
                    ElseIf UCase(subj) = UCase("informe ebitda") Then
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeEbitda', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'CUADRANTE SEMANAL TIENDAS
                    ElseIf UCase(subj) = UCase("cuadrante semanal") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreCuadranteSemanal', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INCIDENCIA FICHAJE
                    ElseIf InStr(UCase(subj), "INCIDENCIA FICHAJE") Then
                        preparaCodigoDeAccionFichaje HTMLText, emailDe, empresa
                        GoTo NEXT_MSG
                    'SOLICITUD AUTORIZACION
                    ElseIf InStr(UCase(subj), "SOLICITUD AUTORIZACION PAGO FACTURA") Then
                        preparaCodigoDeAccionAutorizacion HTMLText, emailDe, empresa
                        GoTo NEXT_MSG
                    'RESOLUCIÓN INCIDENCIA
                    ElseIf InStr(UCase(subj), "RE: INCIDENCIA") Then
                        preparaCodigoDeAccionIncidencia HTMLText, emailDe, empresa
                        GoTo NEXT_MSG
                    'INCIDENCIA CERBERO
                    ElseIf InStr(UCase(subj), "INCIDENCIA CERBERO") Then
                        InterpretaIncidenciaCERBERO HTMLText, emailDe, empresa
                        GoTo NEXT_MSG
                    'INFORME BOCADILLOS
                    ElseIf InStr(UCase(subj), "BOCADILLO") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeBocadillos', 0, '" & Replace(Replace(subj, "_", " ") & "','", "—", "--") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    'INFORME PERSONAL
                    ElseIf Left(UCase(subj), 16) = UCase("informe personal") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformePersonal', 0, '" & Replace(Replace(subj, "_", " ") & "','", "—", "--") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    ElseIf Left(UCase(subj), 13) = UCase("veure tiquets") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreVeureTiquets', 0, '" & Replace(Replace(subj, "_", " ") & "','", "—", "--") & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    ElseIf UCase(Left(subj, 8)) = UCase("Resultat") Then
                        ExecutaComandaSql "use  " & empresa
                        Set rs = Db.OpenResultset("Select distinct Botiga from [" & NomTaulaVentas(Now) & "] where day(data) = " & Day(Now))
                        While Not rs.EOF
                            EnviaEmailAdjunto emailDe, "Excel Resultat ", EnviarResultatBotiga(rs("Botiga"))
                            rs.MoveNext
                        Wend
                        ExecutaComandaSql "use Hit"
                        GoTo NEXT_MSG
                    ElseIf UCase(Left(subj, 8)) = UCase("Registro") Then
                        sf_enviarMail User, emailDe, "Registro", "Formato de registro<Br>  Mensaje : Registro<Br>Texto de mensaje <Br><Br>Tienda: 'Texto tienda'<Br>Texto: 'Mensaje en una sola linea' <Br> Adjuntar ficheros a registrar ;)", "", ""
                                        
                    ElseIf TeElTag(subj, "Numero Serie") Then
                        AsignaFacturaNumSerie Text, empresa
                        sf_enviarMail User, emailDe, "Registro", "Formato de registro<Br>  Mensaje : Registro<Br>Texto de mensaje <Br><Br>Tienda: 'Texto tienda'<Br>Texto: 'Mensaje en una sola linea' <Br> Adjuntar ficheros a registrar ;)", "", ""
                        GoTo NEXT_MSG
                    'FORZAR TRASPASO CAJAS DD/MM/YYYY <CÓDIGO TIENDA>
                    'ElseIf UCase(Left(subj, 15)) = UCase("FORZAR TRASPASO") Then
                    '    ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreForzarTraspaso', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                    '    GoTo NEXT_MSG
                    
                    'RESUMEN HORAS
                    ElseIf UCase(Left(subj, 13)) = UCase("resumen horas") Or UCase(Left(subj, 13)) = UCase("resumen_horas") Then
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreResumenHoras', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    
                    Else
                        ExecutaComandaSql "use  " & empresa
                        ExecutaComandaSql "Insert into feinesAFer (Id, Tipus, Ciclica, Param1, Param2, Param3, Param4, Param5, tmStmp) values (newid(), 'SecreInformeHelp', 0, '" & subj & "','" & emailDe & "', '" & empresa & "', '', '', getdate())"
                        GoTo NEXT_MSG
                    End If
                    
                    'If UCase(Left(subj, 3)) = UCase("Tel") Then
                    '    NomUsu = Trim(Right(subj, Len(subj) - InStr(subj, " ")))
                    '    Dim TelRandom As String
                    '    TelRandom = Right("3" & (Rnd * 10000000), 6)
                    '    ExecutaComandaSql "Delete hit.dbo.Telefonos where idt = '" & NomUsu & "'"
                    '    ExecutaComandaSql "Insert Into hit.dbo.Telefonos (idt,cliente,actdata,actcodi,actestat,domini) Values ('" & NomUsu & "','" & empresa & "',GETDATE(),'" & TelRandom & "','PendentActivacio','') "
                    '    ExecutaComandaSql "Insert Into " & empresa & ".dbo.Recurssos (idt,cliente,actdata,actcodi,actestat,domini) Values ('" & NomUsu & "','" & empresa & "',GETDATE(),'" & TelRandom & "','PendentActivacio','') "
                    '    sf_enviarMail "Secrehit@gmail.com", emailDe, "Telefono : " & NomUsu & " Asignado.", "Desde el telefon " & NomUsu & " llame al 937160210 i cuando se lo pidan teclee el codigo : <Br> " & TelRandom & Chr(13) & Chr(10) & " <Br><Br><Br> Atentamente Joana.", "", ""
                    'End If
                    
                    
                'HAY DOCUMENTO ADJUNTO --------------------------------------------------------------------------------------------------------------------
                'Else
                
                If frmSplash.Pop3Message.Attachments.Count > 0 Then
                    extF = "Jpg"
                    If InStr(frmSplash.Pop3Message.Attachments(0).Name, ".") > 0 Then extF = Right(frmSplash.Pop3Message.Attachments(0).Name, Len(frmSplash.Pop3Message.Attachments(0).Name) - InStr(frmSplash.Pop3Message.Attachments(0).Name, "."))
                    mimeF = "image/jpeg"
                    NomAtt = "AttEmail." & extF
                    MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                    frmSplash.Pop3Message.Attachments(0).Save "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                    Informa "Rebut File " & frmSplash.Pop3Message.Attachments(0).Name
                    t = ""
                    iD = ""
                    subj = frmSplash.Pop3Message.Subject
                    
                    If UCase(Left(subj, 3)) = "ART" Then
                        t = "articles"
                        Set rs = Db.OpenResultset("Select codi,Nom From  " & empresa & ".dbo.articles where nom like '%" & Trim(Right(subj, Len(subj) - InStr(subj, " "))) & "%' ")
                        If rs.EOF Then Set rs = Db.OpenResultset("select a.codi codi , a.nom nom  from  " & empresa & ".dbo.articles a join  " & empresa & ".dbo.articlespropietats p on a.codi = p.codiarticle where variable = 'CODI_PROD'  and valor = '" & Trim(Right(subj, Len(subj) - InStr(subj, " "))) & "' ")
                        
                        If Not rs.EOF Then
                            iD = rs(0)
                            NomUsu = rs("nom")
                        End If
                    End If
                    
                    If UCase(Left(subj, 3)) = "TRE" Or UCase(Left(subj, 3)) = "DEP" Then
                        t = "dependentes"
                        Dim NomDep As String
                        
                        NomDep = Trim(Right(subj, Len(subj) - InStr(subj, " ")))
                        Set rs = Db.OpenResultset("Select codi,nom From " & empresa & ".dbo.dependentes where nom like '%" & NomDep & "%' ")
                        If Not rs.EOF Then
                            iD = rs("Codi")
                            NomUsu = rs("nom")
                        Else
                            ExecutaComandaSql "insert into " & empresa & ".dbo.dependentes  Select max(codi)+1 as codi ,'" & NomDep & "' as Nom , '" & NomDep & "' as memo, '' as telefon , '' as [adreça] , '' as icona, 0 as [hi editem horaris] , '' tid from " & empresa & ".dbo.dependentes "
                            Set rs = Db.OpenResultset("Select codi,nom From " & empresa & ".dbo.dependentes where nom like '%" & NomDep & "%' ")
                            If Not rs.EOF Then
                                iD = rs("Codi")
                                NomUsu = rs("nom")
                            End If
                        End If
                    End If
                
                    If UCase(Left(Trim(subj), 3)) = "FAC" Then
                        t = "Factura"
                        InterpretaFacturaPdf "\Facturacion\ElForn\file\tmp\" & NomAtt, empresa, emailDe
                    End If
                
                    If UCase(Left(subj, 3)) = "INC" Then
                        t = "Incidencias"
                        If InStr(Trim(Right(subj, Len(subj) - InStr(subj, " "))), " ") > 0 Then
                            If IsNumeric(Left(Trim(Right(subj, Len(subj) - InStr(subj, " "))), InStr(Trim(Right(subj, Len(subj) - InStr(subj, " "))), " "))) Then
                                Set rs = Db.OpenResultset("Select id From " & empresa & ".dbo.incidencias where id = '" & Left(Trim(Right(subj, Len(subj) - InStr(subj, " "))), InStr(Trim(Right(subj, Len(subj) - InStr(subj, " "))), " ")) & "' ")
                            Else
                                Set rs = Db.OpenResultset("Select id From " & empresa & ".dbo.incidencias where id = '-9999' ")
                            End If
                        Else
                            If IsNumeric(Trim(Right(subj, Len(subj) - InStr(subj, " ")))) Then
                                Set rs = Db.OpenResultset("Select id From " & empresa & ".dbo.incidencias where id = '" & Trim(Right(subj, Len(subj) - InStr(subj, " "))) & "' ")
                            Else
                                Set rs = Db.OpenResultset("Select id From " & empresa & ".dbo.incidencias where id = '-9999' ")
                            End If
                        End If
                        If Not rs.EOF Then
                            ruta1 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                            If InStr(Trim(Right(subj, Len(subj) - InStr(subj, " "))), " ") > 0 Then
                                salvar ruta1, "INC_" & Left(Trim(Right(subj, Len(subj) - InStr(subj, " "))), InStr(Trim(Right(subj, Len(subj) - InStr(subj, " "))), " ")), extF, mimeF, "INC_" & Trim(Right(subj, Len(subj) - InStr(subj, " "))), NomUsu, n, t, iD, empresa
                            Else
                                salvar ruta1, "INC_" & Trim(Right(subj, Len(subj) - InStr(subj, " "))), extF, mimeF, "INC_" & Trim(Right(subj, Len(subj) - InStr(subj, " "))), NomUsu, n, t, iD, empresa
                            End If
                            MyKill CStr(ruta1)
                            EnviaEmail emailDe, "Foto de incidencia " & rs("id") & " Actualizada"
                        Else
                            'If Not IsNumeric(nInc) Then
                                EnviaEmail emailDe, "Error recibiendo mail no encuentro incidencia, (" & subj & "). Formato INC [ESPACIO] NUMERO DE INCIDENCIA"
                            'End If
                        End If
                    ElseIf InStr(UCase(subj), UCase("incidencia:")) Then
                        Dim nInc As String
                    
                        ExecutaComandaSql "use  Fac_hitrs"
                        nInc = Split(Split(subj, "Incidencia: ")(1), " ")(0)
                    
                        ExecutaComandaSql "UPDATE incidencias SET Estado = 'Resuelta' WHERE id = '" & nInc & "'"
                        ExecutaComandaSql "UPDATE incidencias SET incidencia = incidencia + ' <br> FastByte: ok' WHERE id = '" & nInc & "'"
                    End If
                
                    'If UCase(Left(subj, 3)) = "XML" Then
                    '    t = "Factura"
                    '    InterpretaFacturaXml "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt, empresa, emailDe
                    'End If
            
                    If TeElTag(subj, "Numero Serie") Then
                        resposta = AsignaFacturaNumSerie(Text, empresa)
                        sf_enviarMail User, emailDe, "Numero Serie", resposta, "", ""
                    End If
                
                    If t <> "" And iD <> "" Then
                        URL = "http://www.gestiondelatienda.com/facturacion/elforn/file/tmp/uploadFoto.php?"
                        URL = URL & "rutaImg=" & NomAtt
                        resposta = llegeigHtml(URL)
                        aResposta = Split(resposta, "|")
                        If UBound(aResposta) > 0 Then
                            'Foto
                            ruta1 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(0)
                            salvar ruta1, "ORIGINAL", extF, mimeF, "Foto original de " & NomUsu, NomUsu, n, t, iD, empresa
                            MyKill CStr(ruta1)
                            'Foto Screen
                            ruta2 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(1)
                            salvar ruta2, "SCREEN", extF, mimeF, "Foto pantalla de " & NomUsu, NomUsu, n, t, iD, empresa
                            MyKill CStr(ruta2)
                            'ICO
                            ruta3 = "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & aResposta(2)
                            IdFoto = salvar(ruta3, "ICO", extF, mimeF, "Foto TPV de " & NomUsu, NomUsu, n, t, iD, empresa)
                            'Foto TPV
                            IdFoto = salvar(ruta1, "TPV", extF, mimeF, "Foto TPV de " & NomUsu, NomUsu, n, t, iD, empresa)
                            MyKill CStr(ruta3)
                            EnviaEmail emailDe, "Foto de " & NomUsu & " Actualizada"
                        Else
                            EnviaEmail emailDe, "Error de imatge " & resposta & " ho sento :("
                        End If
                        'Pujar imatge  a botigues
                        rec ("INSERT into missatgesaenviar (Tipus,Param) values ('Imatges" & t & "','') ")
                    Else
                        If t <> "Incidencias" And t <> "Factura" Then
                            EnviaEmail emailDe, "No Se para quien es la foto, (" & subj & "). Formato ART i nom o codi article o DEP i part del nom de la dependenta"
                        End If
                    End If
                    MyKill "C:\Web\gdt\Facturacion\ElForn\file\tmp\" & NomAtt
                End If
            End If
        End If
     
NEXT_MSG:
        If Borrar Then
            frmSplash.POP3.Messages(i).MarkDelete = True
            UnDeBorrat = True
        End If
        If i > 25 Then i = frmSplash.POP3.Messages.Count - 1  ' de 100 en 100
    Next
    
    frmSplash.POP3.Disconnect
    H1 = Now
    While frmSplash.POP3.State = WODPOP3COMLib.StatesEnum.Connected
        DoEventsSleep
        If Now > DateAdd("s", 90, H1) Then Exit Sub
    Wend
    
    If UnDeBorrat Then
        H1 = Now
        While Now <= DateAdd("s", 4, H1)
            DoEventsSleep
        Wend
    End If
    
    Exit Sub
    
nor:
'    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR RevisaEmailDe [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Err.Description, "", ""
    
End Sub




Sub SecreInformeHelp(subj As String, emailDe As String, empresa As String)
    Dim co As String
    
On Error GoTo nor

    co = "<H2>INFORMES</H2>"
        
    co = co & "<H3>Informe de Vendes</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Informe ventas</B></DD>"
    co = co & "<DD><B>Informe ventas dd/mm/yyyy</B></DD>"
    co = co & "<DD><B>Informe ventas SEMANA S</B></DD>"
    co = co & "<BR>"
    co = co & "<DD>La data i la setmana són opcionals</DD>"
    co = co & "<DD>Si posem SEMANA ens envia l'informe de vendes de la setmana S de l'any actual.</DD>"
    co = co & "<DD>Si no posem ni data, ni setmana, ens envia l'informe d'avui.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>El poden demanar, vía email, els usuaris Gerent i Ajudant de gerent. Rebran un email amb informació de totes les supervisores i botigues.<DD>"
    co = co & "<DD>El poden demanar, vía email, les supervisores. Rebran un email amb informació només de les botigues que supervisen.</DD>"
    co = co & "<DD>El poden demanar, vía email, els usuaris franquicia. Rebran informació només de les botigues que supervisen.</DD>"
    co = co & "<DD>Per demanar informes de vendes l'usuari ha d'estar donat d'alta a la ""Secre"".</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>Conté informació de vendes, compres i fitxatges organitzat per supervisores i botigues<DD>"
    co = co & "<DD>"
    co = co & "<h4><I>SUPERVISORA</I></h4>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><Td rowspan='2'><b>Botiga</b></Td><Td align='center' colspan='3'><b>Vendes</b></Td><Td rowspan='2'><b>Clients</b></Td><Td rowspan='2'><b>Tiquet Mig</b></Td><Td colspan='3' align='center'><b>Vendes Any anterior</b></Td><td rowspan='2' align='center'><b>Clients <br>Any anterior</b></td><td rowspan='2' align='center'><b>Tiquet mig <br>Any anterior</b></td><Td rowspan='2'><b>Inc</b></Td><Td colspan='2'><b>Compres</b></Td><Td colspan='2'><b>Devolucions</b></Td><Td colspan='2'><b>Horas</b></Td></Tr>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><td><b>Total</b></Td><td><b>%</b></Td><td><b>Total</b></Td><td><b>%</b></Td><td><b>Total</b></Td><td><b>%</b></Td></Tr>"
    co = co & "<tr><td><b><I>Nom botiga</I></b></td><td><I>Vendes matí</I></td><td><I>Vendes tarda</I></td><td><I>Vendes matí+tarda</I></td><td><I>Número de cliens</I></td><td><I>Tiquet mig<br>Total vendes/N. Clients</I></td>"
    co = co & "<td><I>Vendes matí<br>Any anterior</I></td><td><I>Vendes tarda<br>Any anterior</I></td><td><I>Vendes matí+tarda<br>Any anterior</I></td><td><I>Número de cliens<br>Any anterior</I></td><td><I>Tiquet mig<br>Total vendes/N. Clients<br>Any anterior</I></td>"
    co = co & "<td><I><font color='green'>% Increment vendes</font><br><font color='red'>% Descens vendes</font></I></td><td><I>Import de compres</I></td><td><I>Compres / Vendes sense IVA</I></td><td><I>Import devolucions<br><font color='red'>més de 160 o menor de 85</font></I></td><td><I>Devolucions / Total vendes</I></td>"
    co = co & "<td><I>Total hores dependentes</I></td><td><I>Total hores / Vendes sense IVA</I></td>"
    co = co & "</tr>"
    co = co & "</Table>"
    co = co & "</DD>"
    co = co & "<BR><DD><I>Listat del detall de hores per dependenta a cada botiga</I></DD>"
    co = co & "</DL>"
    
    co = co & "<H3>Informe Supervisores</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Informe supervisora</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>El poden demanar, vía email, els usuaris Gerent i Ajudant de gerent. Rebran un email per cada supervisora.<DD>"
    co = co & "<DD>El poden demanar, vía email, les supervisores. Rebran un email amb informació només de les botigues que supervisen.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>Conté informació de vendes, compres i hores per botigues i acumulats semanals<DD>"
    co = co & "<h4><I>SUPERVISORA</I></h4>"
    co = co & "<DD>Resum de dades de totes les botigues<DD>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><td colspan='19'><b>Data (Ahir)</b><I> Resum de dades de totes les botiges</I></td></tr>"
    co = co & "<Tr bgColor='#DADADA'><Td rowspan='2'><b>Botiga</b></Td><Td align='center' colspan='4'><b>Vendes</b></Td><Td align='center' colspan='2'><b>Diferència Clients</b></Td><Td align='center' colspan='2'><b>Tiquet mig</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td align='center' colspan='2'><b>Devolucions</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td align='center' colspan='4'><b>Hores</b></Td></Tr>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Dif dia</b></Td><td><b>Dif Acumulat</b></Td><Td><b>Any anterior<br>Any actual</b></td><td><b>Acumulat</b></Td><Td><b>Dia</b></td><td><b>Acumulat</b></Td><Td><b>Dia</b></td><td><b>% Acumulat</b></Td><Td><b>Dia</b></td><td><b>Acumulat</b></Td><td><b>Dia</b></Td><td><b>%</b></Td><td><b>Dia</b></Td><td><b>%</b></Td><td><b>Dif. Acumulat</b></Td><td><b>%</b></Td></Tr>"
    co = co & "<Tr><td><b><I>Nom botiga</I></b></td><td><I>Vendes matí</I></td><td><I>Vendes tarda</I></td><td><I>Vendes vs Previsió</I></td><td><I>Vendes vs Previsió<br>Acumulat setmanal</I></td><td><I>Dif. clients<BR>Dia any anterior</I></td><td><I>Dif. clients<BR>Acumulat setmanal</I></td>"
    co = co & "<td><I>Tiquet mig vs Límit tiquet mig</I></td><td><I>Tiquet mig vs Límit tiquet mig<br>Acumulat setmanal</I></td><td><I>Import Compres<br>Acumulat setmanal</I></td><td><I>% compres vs vendes<br>Acumulat setmanal</I></td><td><I>Import devolucions del dia</I></td><td><I>Import devolucions<br>Acumulat setmanal</I></td>"
    co = co & "<td><I>Import de compres</I></td><td><I>Compres / Vendes sense IVA</I></td>    "
    co = co & "<td><I>Total hores dia</I></td><td><I>% Total hores / Vendes sense IVA</I></td><td><I>Total hores dia<br>Acumulat setmanal</I></td><td><I>% Total hores / Vendes sense IVA<br>Acumulat setmanal</I></td>"
    co = co & "</tr>"
    co = co & "</Table>"
    co = co & "<BR>"
    co = co & "<DD>Detall setmanal de cada botiga<DD>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><td colspan='19'><b>Botiga</b><I> Detall setmanal (desde ahir fins dilluns anterior) dia a dia de cada botiga</I></td></tr>"
    co = co & "<Tr bgColor='#DADADA'><Td rowspan='2'><b>Data</b></Td><Td align='center' colspan='4'><b>Vendes</b></Td><Td align='center' colspan='2'><b>Diferència Clients</b></Td><Td align='center' colspan='2'><b>Tiquet mig</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td align='center' colspan='2'><b>Devolucions</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td align='center' colspan='4'><b>Hores</b></Td></Tr>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Dif dia</b></Td><td><b>Dif Acumulat</b></Td><Td><b>Any anterior<br>Any actual</b></td><td><b>Acumulat</b></Td><Td><b>Dia</b></td><td><b>Acumulat</b></Td><Td><b>Dia</b></td><td><b>% Acumulat</b></Td><Td><b>Dia</b></td><td><b>Acumulat</b></Td><td><b>Dia</b></Td><td><b>%</b></Td><td><b>Dia</b></Td><td><b>%</b></Td><td><b>Dif. Acumulat</b></Td><td><b>%</b></Td></Tr>"
    co = co & "<Tr><td><b><I>Data</I></b></td><td><I>Vendes matí</I></td><td><I>Vendes tarda</I></td><td><I>Vendes vs Previsió</I></td><td><I>Vendes vs Previsió<br>Acumulat setmanal</I></td><td><I>Dif. clients<BR>Dia any anterior</I></td><td><I>Dif. clients<BR>Acumulat setmanal</I></td>"
    co = co & "<td><I>Tiquet mig vs Límit tiquet mig</I></td><td><I>Tiquet mig vs Límit tiquet mig<br>Acumulat setmanal</I></td><td><I>Import Compres<br>Acumulat setmanal</I></td><td><I>% compres vs vendes<br>Acumulat setmanal</I></td><td><I>Import devolucions del dia</I></td><td><I>Import devolucions<br>Acumulat setmanal</I></td>"
    co = co & "<td><I>Import de compres</I></td><td><I>Compres / Vendes sense IVA</I></td>    "
    co = co & "<td><I>Total hores dia</I></td><td><I>% Total hores / Vendes sense IVA</I></td><td><I>Total hores dia<br>Acumulat setmanal</I></td><td><I>% Total hores / Vendes sense IVA<br>Acumulat setmanal</I></td>"
    co = co & "</tr>"
    co = co & "</Table>"
    
    co = co & "</DL>"

    
    co = co & "<H3>Informe Coordinadores</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Informe coordinadora</B></DD>"
    co = co & "<DD><B>Informe coordinadora <I>Nom botiga</I></B></DD>"
    co = co & "<BR>"
    co = co & "<DD>Si no indiquem botiga ens envia un correu per cada botiga.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>El poden demanar, vía email, els usuaris Gerent i Ajudant de gerent. Rebran un email per cada coordinadora.</DD>"
    co = co & "<DD>El poden demanar, vía email, les coordinadores. Rebran un email amb informació només de les botigues en les que han fitxat.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"

    
    co = co & "<H3>Informe de Comanda Setmanal</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Pedido semanal <I>Viatge</I></B></DD>"
    co = co & "<BR>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>El poden demanar, vía email, les supervisores.<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>Informa de la comanda (graella) de la setmana actual, dels productes introduïts al viatge demanat, per totes les botigues que supervisa, la supervisora que el demana.</DD>"
    co = co & "</DL>"

    
    co = co & "<H3>Informe de Productes</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Informe productos</B></DD>"
    co = co & "<DD><B>Informe productos dd/mm/yyyy</B></DD>"
    co = co & "<DD><B>Informe productos SEMANA S</B></DD>"
    co = co & "<BR>"
    co = co & "<DD>La data i la setmana són opcionals</DD>"
    co = co & "<DD>Si posem SEMANA ens envia l'informe de vendes de la setmana S de l'any actual.</DD>"
    co = co & "<DD>Si no posem ni data, ni setmana, ens envia l'informe d'avui.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>El poden demanar, vía email, els usuaris Gerent, Ajudant de gerent i Supervisora.</DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>Informe dels 10 productes més venuts a cada botiga. Si l'usuari és Gerent o Ajudant de Gerent informa del 10 productes més venuts, en global, a totes les botigues propies i el detall per botiga. Les supervisores només podran veure el detall de les botigues que supervisen.</DD>"
    co = co & "</DL>"

        
    co = co & "<H3>Informe de Uso de Productos</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>Com el demano?</H4></DT>"
    co = co & "<DD><B>Listado productos en uso</B></DD>"
    co = co & "<BR>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>Tots els usuaris que tinguin registrat el seu email.<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>LListat dels productes que s'han fet servir els últims 12 mesos (Producte, Familia, Preu, Iva, Última venda, Última factura, Últim albarà<DD>"
    co = co & "<DD>Productes obsolets. Fa més de 100 dies que no es fan servir (no s'han venut per botiga, ni per albarà).</DD>"
    co = co & "</DL>"

        
    co = co & "<H3>Informe de Previsiones</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Como lo pido?</H4></DT>"
    co = co & "<DD><B>xx</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>xx<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"

        
    co = co & "<H3>Informe de Incidencias</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Como lo pido?</H4></DT>"
    co = co & "<DD><B>xx</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Qui el pot demanar?</H4></DT>"
    co = co & "<DD>xx<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>Quina informació conté?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"

    
    co = co & "<H3>Informe de Masas</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Como lo pido?</H4></DT>"
    co = co & "<DD><B>xx</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Quien lo recibe?</H4></DT>"
    co = co & "<DD>xx<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Que información contiene?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"

    
    co = co & "<H3>Informe de Traspaso a contabilidad</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Como lo pido?</H4></DT>"
    co = co & "<DD><B>xx</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Quien lo recibe?</H4></DT>"
    co = co & "<DD>xx<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Que información contiene?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"

    
    co = co & "<H3>Informe de Cuadrantes (Turnos)</H3>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Como lo pido?</H4></DT>"
    co = co & "<DD><B>Cuadrante Semanal</B></DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Quien lo recibe?</H4></DT>"
    co = co & "<DD>xx<DD>"
    co = co & "</DL>"
    co = co & "<DL>"
    co = co & "<DT><H4>&iquest;Que información contiene?</H4></DT>"
    co = co & "<DD>xxx<DD>"
    co = co & "</DL>"
         
    sf_enviarMail "secrehit@hit.cat", emailDe, UCase(subj) & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    
    Exit Sub
     
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeHelp  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "LO PIDE: [" & emailDe & "]<br>" & err.Description, "", ""
End Sub


Function SecreHitEmailEmpresa(User As String) As String
    Dim rs As rdoResultset, rsEmp As rdoResultset
       
    SecreHitEmailEmpresa = ""
    Set rs = Db.OpenResultset("select e.db from hit.dbo.Secretaria s join hit.dbo.web_empreses e on e.db = s.empresa where s.Email = '" & User & "'")
    If Not rs.EOF Then
        If Not IsNull(rs(0)) Then
            SecreHitEmailEmpresa = rs(0)
            Exit Function
        End If
    End If
    
    Set rs = Db.OpenResultset("select * from sys.databases where name like 'Fac_%' and name not like '%bak%'")
    While Not rs.EOF
        If ExisteixTaula(rs("name") & ".dbo.dependentesextes") Then
            Set rsEmp = Db.OpenResultset("select * from [WEB]." & rs("name") & ".dbo.dependentesextes where nom='EMAIL' and upper(valor) like '%" & UCase(User) & "%'")
            If Not rsEmp.EOF Then
                SecreHitEmailEmpresa = rs("name")
                Exit Function
            End If
        End If
        rs.MoveNext
    Wend

    'Set rs = Db.OpenResultset("select * from fac_tena.dbo.dependentesextes where nom='EMAIL' and upper(valor) = upper('" & User & "' )")
    'If Not rs.EOF Then
    '    SecreHitEmailEmpresa = "fac_tena"
    'End If
        

End Function




Sub SecreInformeEbitda(email As String)
    Dim rsEmp As rdoResultset, rsEmpHit As rdoResultset, Rs2 As rdoResultset, rs As rdoResultset
    Dim m As Integer, co As String
    Dim empresaHit As String, empresaBD As String
    Dim meses As Variant
    
    meses = Array("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
                        

    Set rsEmp = Db.OpenResultset("select * from [silema_Ts].sage.dbo.empresas")
    While Not rsEmp.EOF
        empresaHit = "9999"
        Set rsEmpHit = Db.OpenResultset("select * from fac_tena.dbo.constantsEmpresa where valor='" & rsEmp("CifDni") & "'")
        If Not rsEmpHit.EOF Then
            If InStr(rsEmpHit("camp"), "_") Then
                empresaHit = Split(rsEmpHit("camp"), "_")(0)
            Else
                empresaHit = "0"
            End If
            empresaBD = "Fac_tena.dbo."
        Else
            Set rsEmpHit = Db.OpenResultset("select * from fac_Hitrs.dbo.constantsEmpresa where valor='" & rsEmp("CifDni") & "'")
            If Not rsEmpHit.EOF Then
                If InStr(rsEmpHit("camp"), "_") Then
                    empresaHit = Split(rsEmpHit("camp"), "_")(0)
                Else
                    empresaHit = "0"
                End If
                empresaBD = "Fac_Hitrs.dbo."
            End If
        End If
        
        co = "<h3>" & rsEmp("Empresa") & "</h3>"
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Mes</td><td><b>Facturado</b></td><td><b>Pendiente cobro</b></td></tr>"
        
        For m = 0 To 11
            Set rs = Db.OpenResultset("select isnull(sum(ImporteAsiento), 0) Importe from [Silema_Ts].sage.dbo.movimientos where codigoempresa=" & rsEmp("CodigoEmpresa") & "  and codigocuenta like '43%' and cargoAbono='D' and comentario like 'F.num : %' and ejercicio=" & Year(Now()) & " and month(fechaasiento)=" & m + 1)
            Set Rs2 = Db.OpenResultset("select isnull(sum(total), 0) ImportPendent from " & empresaBD & "[" & NomTaulaFacturaIva(CDate("01/" & (m + 1) & "/" & Year(Now()))) & "] fIva left join " & empresaBD & "FacturacioComentaris c on fIva.idfactura = c.idfactura where empresacodi=" & empresaHit & " and c.cobrat = 'N'")

            co = co & "<tr><td>" & meses(m) & "</td><td>" & FormatNumber(rs("Importe"), 2) & "</td><td>" & FormatNumber(Rs2("importPendent"), 2) & "</td></tr>"
        Next
        
        co = co & "</table>"
    
        sf_enviarMail "secrehit@hit.cat", email, "Informe EBITDA " & rsEmp("Empresa") & Format(Now(), "dd/mm/yy"), co, "", ""
        sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe EBITDA " & rsEmp("Empresa") & Format(Now(), "dd/mm/yy"), co, "", ""
        
        rsEmp.MoveNext
    Wend

End Sub

Sub SecreInformeIncidencias(subj As String, emailDe As String, empresa As String)
    Dim cerrada As Integer, resuelta As Integer, curso As Integer, pendiente As Integer, perdida As Integer, Total As Integer
    Dim co As String, sql As String, CoResumen As String
    Dim rs As rdoResultset
    Dim idEmpleado As String
    Dim color As String
    
    On Error GoTo nor
    
    cerrada = 0
    resuelta = 0
    curso = 0
    pendiente = 0
    perdida = 0
    Total = 0
    
    Set rs = Db.OpenResultset("select * from dependentesExtes where nom='EMAIL' and valor='" & emailDe & "'")
    If Not rs.EOF Then
        idEmpleado = rs("id")
        
        co = "<h2>Report Incidencias " & Now() & "<h2>"
                
        'Co = Co & "<table border='1'>"
        'Co = Co & "<tr><td><b>Fecha</td><td><b>Cerrada</b></td><td><b>Resuelta</b></td><td><b>Curso</b></td><td><b>Pendiente</b></td><td><b>Perdida</b></td><td><b>Total</b></td></tr>"
        'sql = "select anyo, mes, dia, isnull([Cerrada], 0) Cerrada, isnull([Resuelta], 0) Resuelta, isnull([Curso], 0) Curso, isnull([Pendiente], 0) Pendiente, isnull([Perdida], 0) Perdida, "
        'sql = sql & "isnull([Cerrada], 0)+isnull([Resuelta], 0)+isnull([Curso], 0)+isnull([Pendiente], 0)+isnull([Perdida], 0) Total "
        'sql = sql & "From "
        'sql = sql & "( "
        'sql = sql & "select  year(timestamp) anyo,  month(timestamp) mes, day(timestamp) dia, estado, count(*) n "
        'sql = sql & "From incidencias "
        'sql = sql & "where (timestamp between dateadd(Day, -7, getdate()) and getdate() or FFinReparacion between dateadd(Day, -7, getdate()) and getdate()) and estado in ('Resuelta','Curso','Perdida','Cerrada','Pendiente') and tecnico = " & idEmpleado & " "
        'sql = sql & "group by year(timestamp),  month(timestamp), day(timestamp), estado) DataTable PIVOT ( SUM(n) for estado in ([Resuelta],[Curso],[Perdida],[Cerrada],[Pendiente])) PivoteTable "
        'sql = sql & "order by anyo, mes, dia"
        'Set Rs = Db.OpenResultset(sql)
        'While Not Rs.EOF
        '    Co = Co & "<Tr>"
        '    Co = Co & "<Td>" & Right("00" & Rs("dia"), 2) & "/" & Right("00" & Rs("mes"), 2) & "/" & Rs("anyo") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Cerrada") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Resuelta") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Curso") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Pendiente") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Perdida") & "</Td>"
        '    Co = Co & "<Td>" & Rs("Total") & "</Td>"
        '    Co = Co & "</Tr>"
   
        '    cerrada = cerrada + Rs("Cerrada")
        '    resuelta = resuelta + Rs("Resuelta")
        '    curso = curso + Rs("Curso")
        '    pendiente = pendiente + Rs("Pendiente")
        '    perdida = perdida + Rs("Perdida")
        '    Total = Total + Rs("Total")
        '
        '    Rs.MoveNext
        'Wend
        'Co = Co & "<Tr>"
        'Co = Co & "<Td>TOTAL</Td>"
        'Co = Co & "<Td>" & cerrada & "</Td>"
        'Co = Co & "<Td>" & resuelta & "</Td>"
        'Co = Co & "<Td>" & curso & "</Td>"
        'Co = Co & "<Td>" & pendiente & "</Td>"
        'Co = Co & "<Td>" & perdida & "</Td>"
        'Co = Co & "<Td>" & Total & "</Td>"
        'Co = Co & "</Tr>"
        '
        'Co = Co & "</table>"
        'Co = Co & "<BR>"
        
        'POR CLIENTES
        cerrada = 0
        resuelta = 0
        curso = 0
        pendiente = 0
        perdida = 0
        Total = 0
        
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Cliente</td><td><b>Cerrada</b></td><td><b>Resuelta</b></td><td><b>Curso</b></td><td><b>Pendiente</b></td><td><b>Perdida</b></td><td><b>Total</b></td></tr>"
        sql = "select nom, isnull([Cerrada], 0) Cerrada, isnull([Resuelta], 0) Resuelta, isnull([Curso], 0) Curso, isnull([Pendiente], 0) Pendiente, isnull([Perdida], 0) Perdida, "
        sql = sql & "isnull([Cerrada], 0)+isnull([Resuelta], 0)+isnull([Curso], 0)+isnull([Pendiente], 0)+isnull([Perdida], 0) Total "
        sql = sql & "From "
        sql = sql & "( "
        sql = sql & "select isnull(c.nom, Icli.nom) nom, estado, count(*) n "
        sql = sql & "From incidencias i "
        sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
        sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
        sql = sql & "where  (timestamp between dateadd(Day, -7, getdate()) and getdate() or FFinReparacion between dateadd(Day, -7, getdate()) and getdate()) and estado in ('Resuelta','Curso','Perdida','Cerrada','Pendiente') and tecnico = " & idEmpleado & " "
        sql = sql & "group by isnull(c.nom, Icli.nom), estado) DataTable PIVOT ( SUM(n) for estado in ([Resuelta],[Curso],[Perdida],[Cerrada],[Pendiente])) PivoteTable "
        sql = sql & "order by nom "
        Set rs = Db.OpenResultset(sql)
        While Not rs.EOF
            co = co & "<Tr>"
            co = co & "<Td>" & rs("nom") & "</Td>"
            co = co & "<Td>" & rs("Cerrada") & "</Td>"
            co = co & "<Td>" & rs("Resuelta") & "</Td>"
            co = co & "<Td>" & rs("Curso") & "</Td>"
            co = co & "<Td>" & rs("Pendiente") & "</Td>"
            co = co & "<Td>" & rs("Perdida") & "</Td>"
            co = co & "<Td>" & rs("Total") & "</Td>"
            co = co & "</Tr>"
    
            cerrada = cerrada + rs("Cerrada")
            resuelta = resuelta + rs("Resuelta")
            curso = curso + rs("Curso")
            pendiente = pendiente + rs("Pendiente")
            perdida = perdida + rs("Perdida")
            Total = Total + rs("Total")
            
            rs.MoveNext
        Wend
        co = co & "<Tr>"
        co = co & "<Td>TOTAL</Td>"
        co = co & "<Td>" & cerrada & "</Td>"
        co = co & "<Td>" & resuelta & "</Td>"
        co = co & "<Td>" & curso & "</Td>"
        co = co & "<Td>" & pendiente & "</Td>"
        co = co & "<Td>" & perdida & "</Td>"
        co = co & "<Td>" & Total & "</Td>"
        co = co & "</Tr>"
            
        co = co & "</table>"
        co = co & "<BR>"
        
                
        'PENDIENTES
        Dim P As Integer
        P = 0
        sql = "select i.id, i.timestamp, isnull(i.incidencia, '') Titulo, isnull(d.nom, '') Responsable, isnull(c.nom, icli.nom) Cliente "
        sql = sql & "from incidencias i "
        sql = sql & "left join dependentes d on i.tecnico=d.codi "
        sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
        sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
        sql = sql & "where estado = 'Pendiente' and tecnico = " & idEmpleado & " "
        sql = sql & "order by timestamp"
        Set rs = Db.OpenResultset(sql)
        
        co = co & "<h3>Pendientes</h3>"
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Código</td><td><b>Fecha</td><td><b>Titulo</b></td><td><b>Responsable</b></td><td><b>Cliente</b></td></tr>"
        While Not rs.EOF
            color = "#FFFFFF"
            If DateDiff("d", rs("timestamp"), Now()) > 15 Then color = "#FF0808"
            co = co & "<tr bgcolor='" & color & "'><td><b>" & rs("id") & "</td><td><b>" & rs("timestamp") & "</td><td><b>" & rs("Titulo") & "</b></td><td><b>" & rs("Responsable") & "</b></td><td><b>" & rs("Cliente") & "</b></td></tr>"
            P = P + 1
            rs.MoveNext
        Wend
        co = co & "</table>"
        
        CoResumen = CoResumen & "<H3>PENDIENTES: " & P & "</H3>"
        
        'CURSO
        P = 0
        sql = "select i.id, i.timestamp, isnull(i.incidencia, '') Titulo, isnull(d.nom, '') Responsable, isnull(c.nom, icli.nom) Cliente "
        sql = sql & "from incidencias i "
        sql = sql & "left join dependentes d on i.tecnico=d.codi "
        sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
        sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
        sql = sql & "where estado = 'Curso' and tecnico = " & idEmpleado & " "
        sql = sql & "order by timestamp"
        Set rs = Db.OpenResultset(sql)
        
        co = co & "<h3>En curso</h3>"
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Código</td><td><b>Fecha</td><td><b>Titulo</b></td><td><b>Responsable</b></td><td><b>Cliente</b></td></tr>"
        While Not rs.EOF
            co = co & "<tr><td><b>" & rs("id") & "</td><td><b>" & rs("timestamp") & "</td><td><b>" & rs("Titulo") & "</b></td><td><b>" & rs("Responsable") & "</b></td><td><b>" & rs("Cliente") & "</b></td></tr>"
            P = P + 1
            rs.MoveNext
        Wend
        co = co & "</table>"
        CoResumen = CoResumen & "<H3>EN CURSO: " & P & "</H3>"
        
        'RESUELTA
        P = 0
        sql = "select i.id, i.timestamp, isnull(i.incidencia, '') Titulo, isnull(d.nom, '') Responsable, isnull(c.nom, icli.nom) Cliente "
        sql = sql & "from incidencias i "
        sql = sql & "left join dependentes d on i.tecnico=d.codi "
        sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
        sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
        sql = sql & "where estado = 'Resuelta' and tecnico = " & idEmpleado & " "
        sql = sql & "order by timestamp"
        Set rs = Db.OpenResultset(sql)
        
        co = co & "<h3>Resuelta</h3>"
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Código</td><td><b>Fecha</td><td><b>Titulo</b></td><td><b>Responsable</b></td><td><b>Cliente</b></td></tr>"
        While Not rs.EOF
            co = co & "<tr><td><b>" & rs("id") & "</td><td><b>" & rs("timestamp") & "</td><td><b>" & rs("Titulo") & "</b></td><td><b>" & rs("Responsable") & "</b></td><td><b>" & rs("Cliente") & "</b></td></tr>"
            P = P + 1
            rs.MoveNext
        Wend
        co = co & "</table>"

        CoResumen = CoResumen & "<H3>RESUELTAS: " & P & "</H3>"
        
        'CERRADAS ÚLTIMA SEMANA
        P = 0
        sql = "select i.id, i.timestamp, i.FFinReparacion, isnull(i.incidencia, '') Titulo, isnull(d.nom, '') Responsable, isnull(c.nom, icli.nom) Cliente "
        sql = sql & "from incidencias i "
        sql = sql & "left join dependentes d on i.tecnico=d.codi "
        sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
        sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
        sql = sql & "where FFinReparacion between dateadd(Day, -7, getdate()) and getdate() and estado in ('Cerrada') and tecnico = " & idEmpleado & " "
        sql = sql & "order by timestamp"
        Set rs = Db.OpenResultset(sql)
        
        co = co & "<h3>Cerradas última semana</h3>"
        co = co & "<table border='1'>"
        co = co & "<tr><td><b>Código</td><td><b>Fecha</td><td><b>Fecha Cierre</td><td><b>Titulo</b></td><td><b>Responsable</b></td><td><b>Cliente</b></td></tr>"
        While Not rs.EOF
            co = co & "<tr><td><b>" & rs("id") & "</td><td><b>" & rs("timestamp") & "</td><td><b>" & rs("FFinReparacion") & "</b></td><td><b>" & rs("Titulo") & "</b></td><td><b>" & rs("Responsable") & "</b></td><td><b>" & rs("Cliente") & "</b></td></tr>"
            P = P + 1
            rs.MoveNext
        Wend
        co = co & "</table>"

        CoResumen = CoResumen & "<H3>CERRADAS EN LA ULTIMA SEMANA: " & P & "</H3>"
                
        
        sf_enviarMail "secrehit@hit.cat", emailDe, "Informe incidencias HIT " & Format(Now(), "hh:nn dd/mm/yy"), CoResumen & co, "", ""
        'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe incidencias HIT " & Format(Now(), "hh:nn dd/mm/yy"), Co, "", ""
    Else
        sf_enviarMail "secrehit@hit.cat", emailDe, "Informe incidencias HIT " & Format(Now(), "hh:nn dd/mm/yy"), "NO ERES USUARIO DE " & empresa, "", ""
    End If
    
    Exit Sub
        
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeIncidencias [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "PETICIÓN [" & subj & "] DE [" & emailDe & "] EMPRESA [" & empresa & "]<br>" & co & "<br>ERROR:" & err.Description, "", """"

    
End Sub

'ELIMINAR: NO SIRVE PARA NADA, AHORA TIENEN SU PROPIO CERBERO
Sub SecreInformeIncidenciasSILEMA(subj As String, emailDe As String, empresa As String)
    'Dim Co As String, Sql As String
    'Dim Rs As rdoResultset, rsSup As rdoResultset
    'Dim idEmpleado As String, client As String, clientSql As String
    'Dim esHit As Boolean
    
    'On Error GoTo nor
    
    'idEmpleado = "152" 'MIQUEL DE SILEMA
    
    
    'esHit = False
    'Set Rs = Db.OpenResultset("select * from WEB.Fac_Hitrs.dbo.dependentesExtes where valor like '%" & emailDe & "%' and nom='EMAIL'")
    'If Not Rs.EOF Then esHit = True
    
    'client = ""
    'clientSql = ""
    'If esHit Then
     '   If UBound(Split(subj, " ")) >= 2 Then
     '       client = Split(subj, " ")(2)
     '       clientSql = " c.nom like '%" & client & "%' "
     '   End If
    'Else
    ''emailDe = "cescuder@silemabcn.com"
     '   Set Rs = Db.OpenResultset("select * from WEB.Fac_Tena.dbo.dependentesExtes where valor like '%" & emailDe & "%' and nom='EMAIL'")
     '   If Not Rs.EOF Then
     '       Sql = "select c.nom "
     '       Sql = Sql & "from WEB.Fac_Tena.dbo.constantsClient cc "
     '       Sql = Sql & "left join WEB.Fac_Tena.dbo.clients c on cc.codi=c.codi "
     '       Sql = Sql & "where variable='SupervisoraCodi' and valor='" & Rs("id") & "' and c.nom is not null "
     '       Set rsSup = Db.OpenResultset(Sql)
     '       If rsSup.EOF Then 'NO ES SUPERVISORA
     '           Exit Sub
     '       Else 'SUPERVISORA
     '           clientSql = "("
     '           While Not rsSup.EOF
     '               If clientSql <> "(" Then clientSql = clientSql & " or "
     '               clientSql = clientSql & " c.nom like '%" & Right(rsSup("nom"), 3) & "%' "
     '               rsSup.MoveNext
     '           Wend
     '           clientSql = clientSql & ") "
     '       End If
     '   Else
     '       Exit Sub
     '   End If
    'End If
        
    'Co = "<h2>Report Incidencias " & Now() & "<h2>"
            
    'PENDIENTES
    'Sql = "select i.estado, i.id, i.timestamp, isnull(i.incidencia, '') Titulo, isnull(d.nom, '') Responsable, isnull(c.nom, '') Cliente "
    'Sql = Sql & "from WEB.Fac_Hitrs.dbo.incidencias i "
    'Sql = Sql & "left join WEB.Fac_Hitrs.dbo.dependentes d on i.tecnico=d.codi "
    'Sql = Sql & "left join WEB.Fac_Hitrs.dbo.clients c on i.cliente=c.codi "
    'Sql = Sql & "where estado in ('Pendiente', 'Curso') and tecnico = " & idEmpleado & " "
    'If clientSql <> "" Then Sql = Sql & " and " & clientSql
    'Sql = Sql & "order by timestamp"
    'Set Rs = Db.OpenResultset(Sql)
    
    'Co = Co & "<h3>Pendientes</h3>"
    'Co = Co & "<table border='1'>"
    'Co = Co & "<tr><td><b>Estado</td><td><b>Código</td><td><b>Fecha</td><td><b>Incidencia</b></td><td><b>Responsable</b></td><td><b>Cliente</b></td></tr>"
    'While Not Rs.EOF
    '    Co = Co & "<tr><td><b>" & Rs("estado") & "</td><td><b>" & Rs("id") & "</td><td><b>" & Rs("timestamp") & "</td><td><b>" & Rs("Titulo") & "</b></td><td><b>" & Rs("Responsable") & "</b></td><td><b>" & Rs("Cliente") & "</b></td></tr>"
    '    Rs.MoveNext
    'Wend
    'Co = Co & "</table>"
        
    'sf_enviarMail "secrehit@hit.cat", emailDe, "Informe incidencias SILEMA " & Format(Now(), "hh:nn dd/mm/yy"), Co, "", ""
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe incidencias HIT " & Format(Now(), "hh:nn dd/mm/yy"), Co, "", ""
    
    'Exit Sub
        
'nor:
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeIncidenciasSILEMA [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "PETICIÓN [" & subj & "] DE [" & emailDe & "] EMPRESA [" & Empresa & "]<br>" & Co & "<br>ERROR:" & err.Description, "", """"

    
End Sub


Sub SecreInformeMasas(subj As String, emailDe As String, empresa As String)
    Dim co As String, sql As String
    Dim rs As rdoResultset
    Dim equip As String
    Dim diasMenos As Integer
    Dim fecha As Date
                        
    On Error GoTo errDias
    diasMenos = Split(subj, " ")(2)
    GoTo okDias
errDias:
    diasMenos = 0
okDias:

    If diasMenos >= 1 Then
        fecha = DateAdd("d", diasMenos * (-1), Now())
    Else
        fecha = Now()
    End If

    co = "<h2>Report Masas " & fecha & "<h2>"
    sql = "select Tmst, d.nom, Equip, Masa, Accio, Aux1 "
    'sql = sql & "from [ProduccioMasaLog_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "] p "
    sql = sql & "from [ProduccioMasaLog_2018-09] p "
    sql = sql & "left join dependentes d on p.[user]=d.codi "
    sql = sql & "Where Day(tmSt) = " & Day(fecha) & " "
    sql = sql & "order by equip, tmst"
    Set rs = Db.OpenResultset(sql)
    
    equip = ""
    While Not rs.EOF
        If equip <> rs("Equip") Then
            If equip <> "" Then co = co & "</table><BR>"
            co = co & "<H3>" & rs("Equip") & "</H3>"
            co = co & "<table border='1'>"
            co = co & "<tr><td><b>Tmst</td><td><b>Usuari</td></b><td><b>Masa</b></td><td><b>Accio</b></td><td><b>&nbsp;</b></td></tr>"
            equip = rs("equip")
        End If
        co = co & "<Tr>"
        co = co & "<Td>" & rs("tmst") & "</Td>"
        co = co & "<Td>" & rs("Nom") & "</Td>"
        co = co & "<Td>" & rs("Masa") & "</Td>"
        co = co & "<Td>" & rs("Accio") & "</Td>"
        co = co & "<Td>" & rs("Aux1") & "</Td>"
        co = co & "</Tr>"
        
        rs.MoveNext
    Wend
        
    co = co & "</table>"

    sf_enviarMail "secrehit@hit.cat", emailDe, "Informe Masas " & Format(fecha, "hh:nn dd/mm/yy"), co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe masas " & Format(fecha, "hh:nn dd/mm/yy"), co, "", ""

End Sub

Sub SecreInformePedidoSemanal(subj As String, emailDe As String, empresa As String)
    Dim lunes As Date, domingo As Date, f As Date, rs As rdoResultset, rsServit As rdoResultset, sql As String, co As String
    Dim emaiBot As String, botNom As String, botCodi As String, viaje As String, emailBot As String
    Dim SQL1 As String, sql2 As String, sql3 As String, fecha As Date
    Dim codiDep As String
                            
On Error GoTo ErrData
    
    viaje = Trim(Mid(subj, Len("pedido semanal ")))
    
    lunes = Now()
    While DatePart("w", lunes, vbMonday, vbFirstFullWeek) <> 1
        lunes = DateAdd("d", 1, lunes)
    Wend
    domingo = DateAdd("d", 6, lunes)
    
GoTo OkData
    
ErrData:
    fecha = Now()
    
OkData:

On Error GoTo nor

     Set rs = Db.OpenResultset("select * from dependentesextes where nom='EMAIL' and upper(valor) like '%' + upper('" & emailDe & "') + '%' order by len(valor)")
     If Not rs.EOF Then
         codiDep = rs("id")
         Set rs = Db.OpenResultset("select * from constantsClient where variable='SupervisoraCodi' and valor='" & codiDep & "'")
         If rs.EOF Then 'NO ES SUPERVISORA
             Exit Sub
         Else 'SUPERVISORA
            sql = "Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora , isnull(cc2.valor, '') emailBotiga "
            sql = sql & "from ConstantsClient cc "
            sql = sql & "left join clients c on cc.codi=c.codi "
            sql = sql & "left join constantsClient cc2 on cc2.codi = c.codi and cc2.variable = 'Email' "
            sql = sql & "left join dependentes d on cc.valor = d.codi "
            sql = sql & "where cc.variable = 'SupervisoraCodi' and cc.valor = '" & codiDep & "' "
            sql = sql & "order by c.nom"
         
            Set rs = Db.OpenResultset(sql)
         End If
     End If
     
     While Not rs.EOF
        emailBot = rs("emailBotiga")
        botNom = rs("nom")
        botCodi = rs("codi")
        co = "<h3>Informe pedido semanal " & botNom & " (" & Format(lunes, "dd/mm/yyyy") & " - " & Format(domingo, "dd/mm/yyyy") & ")</h3>"
        co = co & "<TABLE BORDER='1'><TR><TD><B>ARTICLE</B></TD>"
        For f = lunes To domingo
            co = co & "<TD><B>" & Format(f, "dd/mm/yyyy") & "</B></TD>"
        Next
        co = co & "</TR>"
        
        SQL1 = ""
        sql2 = ""
        sql3 = ""
        For f = lunes To domingo
            SQL1 = SQL1 & "isnull([" & Right("0" & Day(f), 2) & "-" & Right("0" & Month(f), 2) & "-" & Year(f) & "], 0) [" & Right("0" & Day(f), 2) & "-" & Right("0" & Month(f), 2) & "-" & Year(f) & "], "
            sql2 = sql2 & "select Codiarticle, QuantitatDemanada, '" & Right("0" & Day(f), 2) & "-" & Right("0" & Month(f), 2) & "-" & Year(f) & "' fecha from [Servit-" & Format(f, "yy-mm-dd") & "] s where client=" & botCodi & " and viatge='" & viaje & "' "
            If f <> domingo Then sql2 = sql2 & " union all "
            sql3 = sql3 & "[" & Right("0" & Day(f), 2) & "-" & Right("0" & Month(f), 2) & "-" & Year(f) & "] "
            If f <> domingo Then sql3 = sql3 & ", "
        Next
        
        sql = "select " & SQL1 & " nom "
        sql = sql & "From "
        sql = sql & "(select a.nom, s.QuantitatDemanada, s.fecha "
        sql = sql & "from ( " & sql2 & ") s "
        sql = sql & "left join articles a on s.Codiarticle=a.codi ) sT "
        sql = sql & "PIVOT (sum(QuantitatDemanada) for fecha in (" & sql3 & ") ) as pT "
        sql = sql & "order by nom"
        Set rsServit = Db.OpenResultset(sql)
        While Not rsServit.EOF
            co = co & "<TR><TD>" & rsServit("nom") & "</TD>"
            For f = lunes To domingo
                co = co & "<TD>" & rsServit(Right("0" & Day(f), 2) & "-" & Right("0" & Month(f), 2) & "-" & Year(f)) & "</TD>"
            Next
            co = co & "</TR>"
            rsServit.MoveNext
        Wend
'select nom, isnull([11-03-19], 0) [11-03-19], isnull([12-03-19], 0) [12-03-19]
'From
'(select a.nom, s.QuantitatDemanada, s.fecha
'from (
'select Codiarticle, QuantitatDemanada, '11-03-19' fecha from [servit-19-03-11] s where client=1543 and viatge='RUSTIC 2'
'Union All
'select Codiarticle, QuantitatDemanada, '12-03-19' fecha from [servit-19-03-12] s where client=1543 and viatge='RUSTIC 2'
') s
'left join articles a on s.Codiarticle=a.codi ) sT
'PIVOT (sum(QuantitatDemanada) for fecha in ([11-03-19], [12-03-19]) ) as pT
'order by nom
        
        co = co & "</TABLE>"
     
        sf_enviarMail "secrehit@hit.cat", emailDe, UCase(subj) & "  [" & Format(Now(), "dd/mm/yy") & "]", co, "", ""
        sf_enviarMail "secrehit@hit.cat", emailBot, UCase(subj) & "  [" & Format(Now(), "dd/mm/yy") & "]", co, "", ""
        sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", UCase(subj) & "  [" & Format(Now(), "dd/mm/yy") & "]", co, "", ""
     
        rs.MoveNext
     Wend
     
     Exit Sub
nor:

End Sub




Sub SecreInformeUsoProductos(subj As String, emailDe As String, empresa As String)
    Dim rsArticles As rdoResultset, rsVenut As rdoResultset, article As String
    Dim co As String, sql As String
    Dim venut As Boolean
    Dim fecha As Date
    Dim i As Integer
    
    On Error GoTo nor
     
    co = ""
    co = co & "<H3>INFORME USO PRODUCTOS</H3>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Producte</b></Td><Td><b>Familia</b></Td><Td><b>Preu</b></Td><Td><b>2on preu</b></Td><Td><b>Iva</b></Td><Td><b>Ultima venta</b></Td><Td><b>Ultima factura</b></Td><Td><b>Ultimo albaran</b></Td></Tr>"
        
    ExecutaComandaSql "drop table LISTADO_PRODUCTOS_VENDIDOS"
    
    sql = "select plu, max(data) data, 'Venta' Tipo into LISTADO_PRODUCTOS_VENDIDOS "
    sql = sql & "from ("
    fecha = Now()
    i = 0
    While DateDiff("m", fecha, Now()) < 12
        If i > 0 Then sql = sql & " union all "
        sql = sql & "select * from [" & NomTaulaVentas(fecha) & "] "
        fecha = DateAdd("m", -1, fecha)
        i = i + 1
    Wend
    sql = sql & ") v "
    sql = sql & "group by plu"
   
    ExecutaComandaSql sql
   
    ExecutaComandaSql "drop table LISTADO_PRODUCTOS_FACTURADOS"
    
    sql = "select producte , max(data) data, 'Factura' Tipo into LISTADO_PRODUCTOS_FACTURADOS "
    sql = sql & "from ("
    fecha = Now()
    i = 0
    While DateDiff("m", fecha, Now()) < 12
        If i > 0 Then sql = sql & " union all "
        sql = sql & "select * from [" & NomTaulaFacturaData(fecha) & "] "
        fecha = DateAdd("m", -1, fecha)
        i = i + 1
    Wend
    sql = sql & ") v "
    sql = sql & "group by producte"
   
    ExecutaComandaSql sql
    
    ExecutaComandaSql "drop table LISTADO_PRODUCTOS_SERVIDOS"

    sql = "select codiarticle, max(data) data, 'Servit' Tipo into LISTADO_PRODUCTOS_SERVIDOS "
    sql = sql & "from ("
    fecha = Now()
    i = 0
    While DateDiff("d", fecha, Now()) < 100
        If i > 0 Then sql = sql & " union all "
        sql = sql & "select distinct codiarticle, '" & Right("0" & Day(fecha), 2) & "/" & Right("0" & Month(fecha), 2) & "/" & Year(fecha) & "' data from [Servit-" & Format(fecha, "yy-mm-dd") & "] "
        fecha = DateAdd("d", -1, fecha)
        i = i + 1
    Wend
    sql = sql & ") s "
    sql = sql & "group by codiarticle "
    
    ExecutaComandaSql sql
   
    sql = "SELECT article, familia, preu, preuMajor, iva, isnull(convert(nvarchar, [Factura], 103), '') Factura, isnull(convert(nvarchar, [Venta], 103), '') Venta, isnull(convert(nvarchar, [Servit], 103), '') albaran "
    sql = sql & "From "
    sql = sql & "(select isnull(a.nom, az.nom) article, isnull(a.familia, az.familia) familia, isnull(a.preu, az.preu) preu, isnull(a.PreuMajor, az.PreuMajor) PreuMajor, t.iva, l.data, l.tipo "
    sql = sql & "from (select * from  LISTADO_PRODUCTOS_VENDIDOS union all select * from LISTADO_PRODUCTOS_FACTURADOS union all select codiArticle plu, data, tipo from LISTADO_PRODUCTOS_SERVIDOS) l "
    sql = sql & "left join articles a on a.codi = l.plu "
    sql = sql & "left join articles_zombis az on az.codi = l.plu "
    sql = sql & "left join TipusIva2012 t on isnull(a.tipoiva, az.tipoiva) = t.tipus "
    sql = sql & ") as DataTable "
    sql = sql & "PIVOT (  max(data)  FOR tipo IN ([Factura], [Venta], [Servit]) "
    sql = sql & ") AS PivotTable "
    sql = sql & "where article is not null "
    sql = sql & "order by article"
    
    Set rsArticles = Db.OpenResultset(sql)
    While Not rsArticles.EOF
        co = co & "<tr><td>" & rsArticles("Article") & "</td><Td>" & rsArticles("Familia") & "</Td><Td>" & rsArticles("Preu") & "</Td><Td>" & rsArticles("PreuMajor") & "</Td><Td>" & rsArticles("Iva") & "</Td><td>" & rsArticles("Venta") & "</td><td>" & rsArticles("Factura") & "</td><td>" & rsArticles("Albaran") & "</td></tr>"
        rsArticles.MoveNext
    Wend
    co = co & "</table>"
    
    co = co & "<h2>PRODUCTOS OBSOLETOS</h2>"
    co = co & "<h3>HACE MAS DE 100 DIAS QUE NO SE USAN</h3>"
    
    sql = "select nom from articles "
    sql = sql & "where codi not in (select plu from  LISTADO_PRODUCTOS_VENDIDOS union all select producte from LISTADO_PRODUCTOS_FACTURADOS union all select codiArticle from LISTADO_PRODUCTOS_SERVIDOS) "
    sql = sql & "and codi not in (select ProdVenut from equivalenciaProductes) and codi not in (select ProdServit from equivalenciaProductes) "
    sql = sql & "order by nom"
    Set rsArticles = Db.OpenResultset(sql)
    While Not rsArticles.EOF
        co = co & rsArticles("nom") & "<br>"
        rsArticles.MoveNext
    Wend

    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""

nor:
End Sub


Sub SecreInformeLicencias(subj As String, emailDe As String, empresa As String)
    Dim rsEmpresas As rdoResultset, rsTocs As rdoResultset, RsAlbarans As rdoResultset, rsRaros As rdoResultset
    Dim co As String, sql As String
    Dim anyo As Integer, mes As Integer, tServit As String
    Dim emp As String, totalEmpresa As Integer, totalTOC As Integer, totalAlbarans As Integer
    
    On Error GoTo nor
     
    co = ""
    co = co & "<H3>INFORME LICENCIAS</H3>"
    
    '--EMPRESAS ACTIVAS
    co = co & "<H3>EMPRESAS ACTIVAS</H3>"
    Set rsEmpresas = Db.OpenResultset("select distinct upper(empresa) empresa from REPASO_CUOTAS order by upper(empresa)")
    While Not rsEmpresas.EOF
        co = co & "<DD>" & rsEmpresas("empresa") & "<br>"
        
        rsEmpresas.MoveNext
    Wend

    anyo = Year(Now())
    mes = Month(Now()) + 1
    If mes = 13 Then
         mes = 1
         anyo = anyo + 1
    End If
    tServit = "SERVIT-" & Right(anyo, 2) & "-" & Right("00" & mes, 2) & "-01"
    If Not ExisteixTaula(tServit) Then CreaTaulaServit2 tServit
   
    '--TOCS
    emp = ""
    totalEmpresa = 0
    totalTOC = 0
    
    sql = "select upper(empresa) Empresa, LICENCIA, c.nom, isnull((select quantitatServida from [" & tServit & "] where client=cc.codi and codiarticle=1068), 0) Alb "
    sql = sql & "from REPASO_CUOTAS rc "
    sql = sql & "left join constantsClient cc on rc.licencia = cc.valor and cc.variable='Ordreruta' "
    sql = sql & "left join clients c on cc.codi = c.codi "
    sql = sql & "where tipo='TOC' "
    sql = sql & "order by upper(empresa)"
    Set rsTocs = Db.OpenResultset(sql)
    
    While Not rsTocs.EOF
        If emp <> rsTocs("empresa") Then
            If emp <> "" Then
                co = co & "<Tr bgColor='#DADADA'><Td colspan='2'><b>Total</b></Td><Td><b>" & totalEmpresa & "</b></Td></Tr>"
                co = co & "</table>"
            End If
            
            co = co & "<h4>" & rsTocs("empresa") & "</h4>"
            co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
            co = co & "<Tr bgColor='#DADADA'><Td><b>Llicencia</b></Td><Td><b>Botiga</b></Td><Td><b>Albarà</b></Td></Tr>"
            totalEmpresa = 0
        End If
        
        co = co & "<Tr><Td><b>" & rsTocs("licencia") & "</b></Td><Td><b>" & rsTocs("nom") & "</b></Td><Td><b>" & rsTocs("Alb") & "</b></Td></Tr>"
        
        totalEmpresa = totalEmpresa + 1
        totalTOC = totalTOC + 1
        emp = rsTocs("empresa")
        
        rsTocs.MoveNext
    Wend

    If emp <> "" Then
        co = co & "<Tr bgColor='#DADADA'><Td colspan='2'><b>Total</b></Td><Td><b>" & totalEmpresa & "</b></Td></Tr>"
        co = co & "</table>"
    End If
    
    co = co & "<h4>TOTAL LLICENCIES DE TOC: " & totalTOC & "</H4>"
    
    totalAlbarans = 0
    Set RsAlbarans = Db.OpenResultset("select sum(quantitatServida) total from [" & tServit & "] where codiarticle=1068")
    If Not RsAlbarans.EOF Then totalAlbarans = RsAlbarans("total")
    co = co & "<h4>TOTAL ALBARANS DE TOC: " & totalAlbarans & "</H4>"
    
    
    co = co & "<H4>REVISAR CUOTAS TOC</H4>"
    sql = "select isnull(c.nom, cz.nom) nom, isnull(c.codi, cz.codi) codi from [" & tServit & "] s left join clients c on s.client=c.codi left join clients_zombis cz on s.client=cz.codi where codiarticle=1068 and client not in "
    sql = sql & "( "
    sql = sql & "select c.codi "
    sql = sql & "from REPASO_CUOTAS rc "
    sql = sql & "left join constantsClient cc on rc.licencia = cc.valor and cc.variable='Ordreruta' "
    sql = sql & "left join clients c on cc.codi = c.codi "
    sql = sql & "Where tipo='TOC' and (select quantitatServida from [" & tServit & "] where client=cc.codi and codiarticle=1068) = 1 ) "
    sql = sql & "order by isnull(c.nom, cz.nom)"
    Set rsRaros = Db.OpenResultset(sql)
    While Not rsRaros.EOF
        co = co & rsRaros("nom") & "(" & rsRaros("codi") & ")<br>"
        rsRaros.MoveNext
    Wend
    
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""

    Exit Sub
    
nor:

    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERRROR: " & subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co & err.Description, "", ""
End Sub



Function TeElTag(eN, El)

    
    TeElTag = InStr(UCase(eN), UCase(El))
    

End Function


