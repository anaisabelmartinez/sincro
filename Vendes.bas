Attribute VB_Name = "Vendes"
Option Explicit

Sub VendesTiquetMig()
    Dim Sql As String
    Dim fecha As Date
    
On Error GoTo nor:
    fecha = DateAdd("d", -1, Now())
    'fecha = CDate("24/06/2019")

    Informa "CALCULANT TIQUET MIG " & Format(fecha, "dd/mm/yyyy"), True
    
    Sql = "insert into [" & TaulaTiquetMig(fecha) & "] (Botiga, Data, Hora, Vendes, Clients, TiquetMig, Tmst) "
    Sql = Sql & "select botiga, convert(datetime, '" & Format(fecha, "dd/mm/yyyy") & "', 103) data, datepart(hour, data) Hora, sum(import) Vendes, count(distinct num_tick) clients, sum(import)/count(distinct num_tick) tiquetMig, getdate() tmst "
    Sql = Sql & "From [" & NomTaulaVentas(fecha) & "] v "
    Sql = Sql & "left join articles a on v.plu=a.codi "
    Sql = Sql & "left join families f3 on a.familia = f3.nom "
    Sql = Sql & "left join families f2 on f3.pare = f2.nom "
    Sql = Sql & "left join families f1 on f2.pare = f1.nom "
    Sql = Sql & "Where Day(data) = " & Day(fecha) & " and f1.nom not like '%diada%' and f2.nom not like '%diada%' and f3.nom not like '%diada%' "
    Sql = Sql & "group by botiga, datepart(hour, data) "
    'sql = sql & "order by botiga, datepart(hour, data)"
    ExecutaComandaSql Sql
    
    Exit Sub
    
nor:

    sf_enviarMail "", "ana@solucionesit365.com", "ERROR VendesTiquetMig: ", Sql & err.Description, "", ""
End Sub


Function getObjetivoTiquetMig(codiBot As Double, fecha As Date) As Double
    Dim tiquetMig As Double
    Dim tiquetMig1 As Double, tiquetMig2 As Double, tiquetMig3 As Double, tiquetMig4 As Double, tiquetMig5 As Double
    Dim rsTM As rdoResultset
    Dim Semana As Integer
    
    tiquetMig = 0
    'MEDIA ÚLTIMAS 5 SEMANAS
    'Le restamos un 5% a las ventas
    
    For Semana = 1 To 5
        Set rsTM = Db.OpenResultset("select isnull((sum(vendes)-(sum(vendes)*0.05))/sum(clients), 4) tM from [" & TaulaTiquetMig(DateAdd("d", -7 * Semana, fecha)) & "] where day(data)=" & Day(DateAdd("d", -7 * Semana, fecha)) & " and botiga=" & codiBot & " ")
        If Not rsTM.EOF Then tiquetMig = tiquetMig + rsTM("tM")
    Next
    rsTM.Close
        
    tiquetMig = tiquetMig / 5
    
    'tiquetMig = 4
    'Set rsTM = Db.OpenResultset("select * from constantsClient where variable='Tiquet_Mig' and codi='" & codiBot & "'")
    'If Not rsTM.EOF Then
    '    If IsNumeric(rsTM("valor")) Then
    '        If tiquetMig > 0 Then tiquetMig = rsTM("valor")
    '    End If
    'End If
    
    
    getObjetivoTiquetMig = tiquetMig
End Function
