Attribute VB_Name = "Secre"

Sub SecreAvisos()
    Dim sql As String, Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, D As Date, Txt As String, Acu As Double

'On Error GoTo nor
     
    If UCase(EmpresaActual) = UCase("panet") Then
    EmpresaActual = EmpresaActual
    End If

    Set Rs = Db.OpenResultset("Select * From " & NomTaulaSecreAvisos & " ")
    While Not Rs.EOF
        Select Case Rs("Tipus")
'select i.incidencia,r.nombre,timestamp from Incidencias i join Recursos r on r.id=i.recurso  where i.estado = 'Resuelta' order by timestamp desc
            Case "SaldoBancari"
                Dim Total As Double
                D = DateAdd("d", -1, Now)
                Set Rs0 = Db.OpenResultset("select Distinct comu_numcuenta from " & DonamNomTaulaNorma43() & "  ")
                Total = 0
                While Not Rs0.EOF
                    Acu = 0
                    Txt = ""
                    Set Rs2 = Db.OpenResultset("select Top 1 * From norma43 where comu_numcuenta = '" & Rs0("comu_numcuenta") & "' order by hit_datavalor desc ,numlineaA desc ")
                    Acu = CDbl(Rs2("comu_saldoInicial")) / 100
                    Set Rs3 = Db.OpenResultset("select Entidad from hit.dbo.codigosentidades where  identidad = '" & Rs2("Comu_Banco") & "' ")
                    If Not Rs3.EOF Then If Not IsNull(Rs3(0)) Then Txt = Txt & " " & Rs3(0)
                    Txt = Txt & " " & Format(Rs2("hit_datavalor"), " dd-mm ")
                    
                    Set Rs2 = Db.OpenResultset("select sum(Hit_importe) T From norma43 where idfichero = '" & Rs2("idfichero") & "' ")
                    Acu = Acu + CDbl(Rs2(0))
                    Txt = Format(Acu, "#,#.00") & " € " & Txt
                    
                    Total = Total + Acu
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('Saldo','" & Rs("usuari") & "','" & Txt & "')"
                    Rs0.MoveNext
                Wend
                If Not Total = Acu Then
                    Txt = " Total : " & Format(Total, "#,#.00") & " €"
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('Saldo','" & Rs("usuari") & "','" & Txt & "')"
                End If
                ExecutaComandaSql "Delete " & NomTaulaSecreAvisos & " Where Id = '" & Rs("Id") & "' "
                
            Case "TrucadesTeves"
                D = DateAdd("d", -1, Now)
                
                If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))
                Set Rs2 = Db.OpenResultset("select i.incidencia,i.contacto,r.nombre,timestamp from Incidencias i join Recursos r on r.id=i.recurso where not i.estado = 'Resuelta' and TimeStamp > convert(datetime,'" & D & "',103) And i.tecnico=" & Rs("usuari") & " order by timestamp desc ")
                While Not Rs2.EOF
                    If D < Rs2("timestamp") Then D = Rs2("timestamp")
                    Txt = ""
                    Txt = Txt & "Trucada : " & Rs2("Nombre")
                    If Rs2("contacto") <> "" Then Txt = Txt & " (" & Rs2("contacto") & ") "
                    Txt = Txt & "; " & Rs2("Incidencia") & "(" & Format(Rs2("TimeStamp"), "hh:nn") & ")"
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "')"
                    Rs2.MoveNext
                Wend
                If IsNull(Rs("Id")) Then
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                Else
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                End If
            
            Case "VendesNONONONONONO"
                D = DateAdd("d", -1, Now)
                
                If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))
                If ExisteixTaula(NomTaulaVentas(Now)) Then
                    Set Rs2 = Db.OpenResultset("select * From [" & NomTaulaVentas(Now) & "] with (nolock) Where data > convert(datetime,'" & D & "',103) order by Data ")
                    While Not Rs2.EOF
                        If D < Rs2("data") Then D = Rs2("data")
                        Txt = ""
                        Txt = Txt & "Venut A " & BotigaCodiNom(Rs2("Botiga")) & " -> " & Rs2("Import") & " € (" & Rs2("Quantitat") & ") " & Trim(ArticleCodiNom(Rs2("plu"))) & Format(Rs2("data"), " (hh:nn)")
                        ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "')"
                        Rs2.MoveNext
                    Wend
                    If IsNull(Rs("Id")) Then
                        ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                    Else
                        ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                    End If
                End If
            
            Case "TrucadesNONONONO"
                D = DateAdd("d", -1, Now)
                If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))
                Set Rs2 = Db.OpenResultset("select i.incidencia,r.nombre,timestamp from " & DonamNomTaulaIncidencias & " i join " & DonamNomTaulaRecursos & " r on r.id=i.recurso  where not i.estado = 'Resuelta' and TimeStamp > convert(datetime,'" & D & "',103) order by timestamp desc ")
                While Not Rs2.EOF
                    If D < Rs2("timestamp") Then D = Rs2("timestamp")
                    Txt = ""
                    Txt = Txt & "Trucada : " & Rs2("Nombre") & "; " & Rs2("Incidencia") & "(" & Format(Rs2("TimeStamp"), "hh:nn") & ")"
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "')"
                    Rs2.MoveNext
                Wend
                If IsNull(Rs("Id")) Then
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                Else
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                End If
            Case "ProduccioNONONONO"
                D = DateAdd("d", -1, Now)
                
                If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))
                Set Rs2 = Db.OpenResultset("select * from produccion where fecha > convert(datetime,'" & D & "',103) order by fecha desc ")
                While Not Rs2.EOF
                    If D < Rs2("fecha") Then D = Rs2("fecha")
                    Txt = ""
                    Txt = Txt & "Fabricats " & Rs2("nCajas") & " " & ArticleCodiNom(Rs2("Plu"))
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "')"
                    Rs2.MoveNext
                Wend
                If IsNull(Rs("Id")) Then
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                Else
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                End If
            Case "ContractesNONONO"
                D = DateAdd("d", -1, Now)
                If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))

                sql = "select distinct codi,case isnull(dependentesextes.valor,'') when '' then GETDATE() "
                sql = sql & "else  convert(smalldatetime,dependentesextes.valor,103) "
                sql = sql & "end Tmst, dependentesextes.nom as nomContracte "
                sql = sql & "from dependentes d3 with (nolock) "
                sql = sql & "left join dependentesextes  with (nolock) on d3.codi=dependentesextes.id and dependentesextes.nom = "
                sql = sql & "(select max(nom) from dependentesextes  with (nolock) where nom like 'DATACONTRACTEFIN%' "
                sql = sql & "and dependentesextes.id = d3.CODI) "
                sql = sql & "left join " & NomTaulaSecreAvisosTxt & " sat on dependentesextes.nom=sat.Lliure2 and d3.CODI=sat.Lliure3 and sat.Tipus = 'Contractes' and sat.Usuari='" & Rs("usuari") & "' "
                sql = sql & "where case isnull(dependentesextes.valor,'') when '' then dateadd(d,16,GETDATE()) "
                sql = sql & "else  convert(smalldatetime,dependentesextes.valor,103) "
                sql = sql & "end  <= dateadd(d,15,convert(smalldatetime,getdate(),103)) and ISNULL(sat.Avisat, '') <>'1' "
                sql = sql & "order by tmst desc"
                Set Rs2 = Db.OpenResultset(sql)
                
                While Not Rs2.EOF
                    If D < Rs2("TmSt") Then D = Rs2("TmSt")
                    Txt = ""
                    Txt = Txt & "Contracte de " & DependentaCodiNom(Rs2("codi")) & " apunt de finalitzar"
                    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1,lliure2,lliure3) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "', '" & Rs2("nomContracte") & "', '" & Rs2("codi") & "')"
                    Rs2.MoveNext
                Wend
                If IsNull(Rs("Id")) Then
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                Else
                    ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                End If
            
            Case "Presencia"
                D = DateAdd("d", -1, Now)
                If ExisteixTaula("produccion") And ExisteixTaula(NomTaulaSecreAvisosTxt) Then
                    If Not IsNull(Rs("Lliure1")) Then D = DateAdd("s", 1, Rs("Lliure1"))
                
                    Set Rs2 = Db.OpenResultset("select * from produccion where fecha > convert(datetime,'" & D & "',103) order by fecha desc ")
                    While Not Rs2.EOF
                        If D < Rs2("fecha") Then D = Rs2("fecha")
                        Txt = ""
                        Txt = Txt & "Fabricats " & Rs2("nCajas") & " " & ArticleCodiNom(Rs2("Plu"))
                        ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('" & Rs("Tipus") & "','" & Rs("usuari") & "','" & Txt & "')"
                        Rs2.MoveNext
                    Wend
                    If IsNull(Rs("Id")) Then
                        ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id is null "
                    Else
                        ExecutaComandaSql "Update SecreAvisos Set Lliure1 = '" & D & "' where id = '" & Rs("Id") & "'"
                    End If
                End If
                
        End Select
        Rs.MoveNext
    Wend
    
nor:



End Sub


