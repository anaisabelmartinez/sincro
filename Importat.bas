Attribute VB_Name = "Importat"
Option Explicit
Dim Q_ArticleCodiGeneticCodi As rdoQuery, Q_ArticleCodiGeneticCodi_2 As rdoQuery
Dim Q_ArticleCodiPvp As rdoQuery
Sub CuantosVendidosDesdePedido(Proveedor As String, ChatUser)
    Dim ComandaProveedor As String, Rs As rdoResultset, BotigaNom, NomArticle, DataUltimaVentaServida, His_FechaUltimaVentaNoEncontrada As Date, His_FechaUltimaDevolucionNoEncontrada As Date, His_Inventados, His_Redondeo, DataUltimaVenta As Date, CalDemanar, Primer, AlgunaComanda, Cua
    
    Set Rs = Db.OpenResultset("Select m.id materia,c.codi botiga,p.codiarticle codiarticle, m.nombre mnombre,c.nom,* from ccmateriasprimas m with (nolock) join ccnombrevalor Pr with (nolock) on Pr.valor = '" & Proveedor & "' And m.id=Pr.id and left(Pr.nombre,13) = 'P_REPOSICION_' join ccnombrevalor b with (nolock) on m.id=b.id and left(b.nombre,11) = 'REPOSICION_' join Clients c with (nolock) on Pr.nombre = 'P_REPOSICION_' + cast(c.codi  as nvarchar) And b.nombre = 'REPOSICION_' + cast(c.codi  as nvarchar) join articlespropietats p with (nolock) on p.variable = 'MatPri' and p.valor = m.id  order by  c.nom,m.nombre ")
    BotigaNom = ""
    While Not Rs.EOF
        If Not BotigaNom = BotigaCodiNom(Rs("Botiga")) Then
            DataUltimaVenta = DataUltimaVentaBusca(Rs("Botiga"), Now)
            BotigaNom = BotigaCodiNom(Rs("Botiga"))
            ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('Pedido','" & ChatUser & "','Botiga : " & BotigaNom & " Ultima Venta : " & DataUltimaVenta & " ' )"
        End If
        
        NomArticle = ArticleCodiNom(Rs("codiarticle"))
        AlgunaComanda = PedidosUltimosDatos(Rs("materia"), Rs("Botiga"), Proveedor, DataUltimaVentaServida, His_FechaUltimaVentaNoEncontrada, His_FechaUltimaDevolucionNoEncontrada, His_Inventados, His_Redondeo)
        CalDemanar = PedidoCuantosFaltan(Rs("codiarticle"), Rs("Botiga"), His_FechaUltimaVentaNoEncontrada, DataUltimaVenta)
        
        Cua = ""
        If Not AlgunaComanda Then Cua = "(Cap Comanda)"
        ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('Pedido','" & ChatUser & "','" & NomArticle & " v: " & CalDemanar & Cua & " ' )"
        
        Rs.MoveNext
    Wend
    ExecutaComandaSql "Insert Into " & NomTaulaSecreAvisosTxt & " (Tipus,usuari,lliure1) Values ('Pedido','" & ChatUser & "','Fi Resum ' )"
End Sub

Function ExisteixTaula(Nomtb As String) As Boolean
   Dim Rs As rdoResultset

On Error GoTo nor

   ExisteixTaula = False
   Set Rs = Db.OpenResultset("Select OBJECT_ID('" & Nomtb & "')")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ExisteixTaula = True
   Rs.Close
nor:

End Function






Function ExisteixCamp(Nomtb As String, NomCamp As String) As Boolean
   Dim Rs As rdoResultset

'   If Not Db.StillConnecting Then ConnectaSqlServer LastServer, LastDatabase
On Error GoTo nor

   ExisteixCamp = False
   Set Rs = Db.OpenResultset("SELECT * FROM INFORMATION_SCHEMA.COLUMNS  WHERE TABLE_NAME = '" & Nomtb & "' AND COLUMN_NAME = '" & NomCamp & "'")
   
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ExisteixCamp = True
   Rs.Close
nor:

End Function






Sub ValidaLlicencia(NumLlicencia)
   Dim NovaEmpresa As String, User As String, Pssw As String, NovaCodiLlicencia  As String, NumTel As String, ErrMiss As String, f
    
   NovaCodiLlicencia = NumLlicencia
   NumTel = "935000123"
   User = "050@alehop"
   Pssw = "gratis"
   If Len(NomServerInternet) > 0 Then NumTel = NomServerInternet
   On Error Resume Next
   MkDir AppPath
   If InStr(AppPath, "mpreses") = 0 Then MkDir AppPath & "\Empreses"
   On Error GoTo 0
   While Not LlicenciaValida(NovaCodiLlicencia, NumTel, User, Pssw, ErrMiss)
      frmSplash.Height = 3900
      frmSplash.Estat.Caption = ErrMiss
      frmSplash.OK = False
      frmSplash.CodiLic = NovaCodiLlicencia
      frmSplash.NumTel = NumTel
      frmSplash.User = User
      frmSplash.Pssw = Pssw
      frmSplash.Command1.Enabled = True
      While Not frmSplash.OK
         My_DoEvents
      Wend
      NovaCodiLlicencia = Trim(frmSplash.CodiLic)
      NumTel = frmSplash.NumTel
      User = frmSplash.User
      Pssw = frmSplash.Pssw
      frmSplash.Command1.Enabled = False
   Wend

   If Not (NovaCodiLlicencia = NumLlicencia) Then
      NumLlicencia = NovaCodiLlicencia
   End If
   
'   frmSplash.Height = 1800
   
   
End Sub


Function NomTaulaData(nom As String) As Date
   Dim P As Integer, p1 As Integer, P2 As Integer
   
   P = InStr(nom, "-")
   If P > 0 Then p1 = InStr(P + 1, nom, "-")
   If p1 > 0 Then P2 = InStr(p1 + 1, nom, "-")
   If P > 0 And p1 > 0 And P2 > 0 Then
      NomTaulaData = DateSerial(Mid(nom, P + 1, p1 - P - 1), Mid(nom, p1 + 1, p1 - P - 1), Mid(nom, P2 + 1, Len(nom) - P2))
   End If
End Function



Sub PauseTrigger(D As Date, Pause As Boolean)
   Dim NomTaula As String, sql As String
 
   NomTaula = "Servit-" & Format(D, "yy-mm-dd")
   If Pause Then
      ExecutaComandaSql "DROP TRIGGER [M_" & NomTaula & "] "
   Else
   sql = "CREATE TRIGGER [M_" & NomTaula & "] ON [" & NomTaula & "] "
   sql = sql & "AFTER INSERT,UPDATE,DELETE AS "
   sql = sql & "Update [" & NomTaula & "] Set [TimeStamp] = GetDate(),    [QuiStamp]  = Host_Name() Where Id In (Select Id From Inserted) "
   sql = sql & "Insert Into ComandesModificades Select Id As Id,GetDate() As [TimeStamp],'" & NomTaula & "' As TaulaOrigen From Inserted "
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from [" & NomTaula & "] Where Id In (Select Id From Inserted)"
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp]+'BORRAT!!!',Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from deleted Where not Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatTornada)  Update [" & NomTaula & "] Set [CitaTornada]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatTornada  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaTornada]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatServida)  Update [" & NomTaula & "] Set [CitaServida]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatServida  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaServida]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatDemanada) Update [" & NomTaula & "] Set [CitaDemanada] = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatDemanada AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaDemanada] AS VarChar(255)) Where Id In (Select Id From Inserted) "
      ExecutaComandaSql sql
   End If

End Sub

Function DonamParam(s As String) As String
   Dim P As Double, p1 As Double, P2 As Double
   
   P = InStr(s, ":")
   p1 = P
   P2 = p1
   While p1 > 0
      p1 = InStr(p1 + 1, s, "]")
      If p1 > 0 Then P2 = p1
   Wend
   
   DonamParam = ""
   If (P2 - P - 1) > 0 Then DonamParam = Mid(s, P + 1, P2 - P - 1)
   
End Function




Function ExecutaQuery(Q As rdoQuery, Optional P0, Optional p1, Optional P2, Optional P3, Optional P4, Optional P5, Optional P6, Optional P7, Optional P8, Optional P9) As Boolean
   
   If Not VarType(P0) = vbError Then Q.rdoParameters(0) = P0
   If Not VarType(p1) = vbError Then Q.rdoParameters(1) = p1
   If Not VarType(P2) = vbError Then Q.rdoParameters(2) = P2
   If Not VarType(P3) = vbError Then Q.rdoParameters(3) = P3
   If Not VarType(P4) = vbError Then Q.rdoParameters(4) = P4
   If Not VarType(P5) = vbError Then Q.rdoParameters(5) = P5
   If Not VarType(P6) = vbError Then Q.rdoParameters(6) = P6
   If Not VarType(P7) = vbError Then Q.rdoParameters(7) = P7
   If Not VarType(P8) = vbError Then Q.rdoParameters(8) = P8
   If Not VarType(P9) = vbError Then Q.rdoParameters(9) = P9
   
   On Error GoTo no_Be
      Q.Execute
   On Error Resume Next
   ExecutaQuery = True
   Exit Function
   
no_Be:
   Debuga Format(Now, "dd hh:mm:ss") & " * (" & Q.sql & ")" & err.Description
   
   ExecutaQuery = False
   
End Function


Function DependentaCodiTid(Cl As Double)
   Static Q As rdoQuery
   Dim Rs As rdoResultset
   
On Error Resume Next
   If Len(Q.sql) = 0 Then
      Set Q = Db.CreateQuery("", "Select Tid From Dependentes Where Codi = ? ")
   End If
   
   Q.rdoParameters(0) = Cl
   Set Rs = Q.OpenResultset
   
   If Rs.EOF Then
      DependentaCodiTid = ""
   Else
      DependentaCodiTid = Rs("Tid")
   End If
   Rs.Close
   
End Function



Function DependentaCodiNom(Cl As Double) As String
   Static Q As rdoQuery
   Dim Rs As rdoResultset
   
On Error Resume Next
   If Len(Q.sql) = 0 Then
      Set Q = Db.CreateQuery("", "Select Nom From Dependentes Where Codi = ? ")
   End If
   
   Q.rdoParameters(0) = Cl
   Set Rs = Q.OpenResultset
   
   If Rs.EOF Then
      DependentaCodiNom = ""
   Else
      DependentaCodiNom = Rs("Nom")
   End If
   Rs.Close
   
End Function


Function EmpresaCodiNom(Cl As Double)
    Dim Rs As rdoResultset, Campnom As String
On Error Resume Next

    If Cl = 0 Then
        Campnom = "CampNom"
    Else
        Campnom = Cl & "_CampNom"
    End If
    Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsempresa where camp = '" & Campnom & "'")
   
    If Rs.EOF Then
       EmpresaCodiNom = ""
    Else
       EmpresaCodiNom = Rs("valor")
    End If
    Rs.Close
   
End Function



Function ClienteGrupo(Cl As Double)
   Static Q As rdoQuery
   Dim Rs As rdoResultset
   
   Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsclient where variable = 'Grup_client' And Codi = " & Cl)
   If Not Rs.EOF Then ClienteGrupo = Rs("valor")
   Rs.Close
   
End Function



Function TePermisPer(s As String) As Boolean
   Dim Rs As rdoResultset
   TePermisPer = True
   
On Error GoTo nor
   
   If ExisteixTaula("QueTinc") Then
      Set Rs = Db.OpenResultset("Select * From  QueTinc Where QueEs = 'Permis' And QuinEs = '" & s & "' ")
      If Rs.EOF Then TePermisPer = False
   End If
   
nor:

End Function



Function AtributCodiPreus(Codi As Long, NoAcabat As Boolean, IncPreu1 As Double, IncPct1 As Double, IncPreu2 As Double, IncPct2 As Double) As Boolean
   Dim sql As String, Codis_Dels_Atributs As rdoResultset
   
   AtributCodiPreus = False
   NoAcabat = False
   IncPreu1 = 0
   IncPct1 = 0
   IncPreu2 = 0
   IncPct2 = 0
   
   If Codi <> 0 Then
      sql = ("Select * From Atributs Where Codi = " & Codi & " ")
      Set Codis_Dels_Atributs = Db.OpenResultset(sql)
      If Not Codis_Dels_Atributs.EOF Then
         NoAcabat = Codis_Dels_Atributs("Es prefabricat")
         If Codis_Dels_Atributs("ModificaPreu1") Then
            IncPreu1 = Codis_Dels_Atributs("Increment_Preu_1")
         Else
            IncPct1 = Codis_Dels_Atributs("Increment_Pct_1")
         End If
         If Codis_Dels_Atributs("ModificaPreu2") Then
            IncPreu2 = Codis_Dels_Atributs("Increment_Preu_2")
         Else
            IncPct2 = Codis_Dels_Atributs("Increment_Pct_2")
         End If
      End If
      AtributCodiPreus = True
   End If

End Function


Function ArticleCodiGeneticCodi(Cl) As Long
   Dim Rs As rdoResultset
   
   ArticleCodiGeneticCodi = 0
On Error Resume Next
   If Len(Q_ArticleCodiGeneticCodi.sql) = 0 Then
      Set Q_ArticleCodiGeneticCodi = Db.CreateQuery("", "  Select Codi From Articles    Where CodiGenetic = ? ")
      Set Q_ArticleCodiGeneticCodi_2 = Db.CreateQuery("", "Select Codi From Memotecnics Where Memotecnic  = ? ")
   End If
Cre:
   Q_ArticleCodiGeneticCodi.rdoParameters(0) = Cl
   Set Rs = Q_ArticleCodiGeneticCodi.OpenResultset
   
   If Not Rs.EOF Then
      ArticleCodiGeneticCodi = Rs("Codi")
   Else
      Q_ArticleCodiGeneticCodi_2.rdoParameters(0) = Cl
      Set Rs = Q_ArticleCodiGeneticCodi_2.OpenResultset
      If Not Rs.EOF Then
         ArticleCodiGeneticCodi = Rs("Codi")
      End If
   End If
   
   Rs.Close
   
End Function

Function ArticleCodiPvp(Codi As Double) As Double
   Dim Rs As rdoResultset
      
   ArticleCodiPvp = 0

On Error Resume Next

   If Len(Q_ArticleCodiPvp.sql) = 0 Then
      Set Q_ArticleCodiPvp = Db.CreateQuery("", "Select Preu From Articles Where Codi = ? ")
   End If
   
   Q_ArticleCodiPvp.rdoParameters(0) = Codi
   Set Rs = Q_ArticleCodiPvp.OpenResultset()
   If Not Rs.EOF Then ArticleCodiPvp = Rs("Preu")
   Rs.Close
   
End Function


