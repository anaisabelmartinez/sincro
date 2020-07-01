Attribute VB_Name = "TaulesBD"
Option Explicit

Function taulaCdpPlanificacion(D As Date) As String
    Dim Str As String, sqlT As String
    Dim lunes As Date

    lunes = D
    While DatePart("w", lunes) <> vbMonday
        lunes = DateAdd("d", -1, lunes)
    Wend

    Str = "cdpPlanificacion_" & Format(lunes, "yyyy_mm_dd")
    If Not ExisteixTaula(Str) Then
        sqlT = "CREATE TABLE [dbo].[" & Str & "] ("
        sqlT = sqlT & "[idPlan] [nvarchar] (255) NULL CONSTRAINT [DF_" & Str & "_idPlan] DEFAULT (newid()),"
        sqlT = sqlT & "[fecha] [datetime] NULL,"
        sqlT = sqlT & "[botiga] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[periode] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[idTurno] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[idEmpleado] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[usuarioModif] [nvarchar](255) NULL, "
        sqlT = sqlT & "[fechaModif] [datetime] NULL CONSTRAINT [DF_" & Str & "_fechaModif] DEFAULT (getdate()),"
        sqlT = sqlT & "[activo] [bit] NULL CONSTRAINT [DF_" & Str & "_activo]  DEFAULT (1)"
        sqlT = sqlT & ") ON [PRIMARY]"

        ExecutaComandaSql sqlT
    End If
                    
    taulaCdpPlanificacion = Str
                   
End Function

Function taulaCdpValidacionHoras(D As Date) As String
    Dim Str As String, sqlT As String

    Str = "cdpValidacionHoras_" & Year(D)

    If Not ExisteixTaula(Str) Then
        sqlT = "CREATE TABLE [dbo].[" & Str & "] ("
        sqlT = sqlT & "[idPlan] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[fecha] [datetime] NULL,"
        sqlT = sqlT & "[botiga] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[dependenta] [nvarchar] (255) NULL,"
        sqlT = sqlT & "[usuarioModif] [nvarchar](255) NULL, "
        sqlT = sqlT & "[fechaModif] [datetime] NULL CONSTRAINT [DF_" & Str & "_fechaModif] DEFAULT (getdate()),"
        sqlT = sqlT & "[validado] [bit] NULL CONSTRAINT [DF_" & Str & "_validado]  DEFAULT (0)"
        sqlT = sqlT & ") ON [PRIMARY]"

        ExecutaComandaSql sqlT
    End If
                    
    taulaCdpValidacionHoras = Str
                   
End Function


Function NomTaulaMovi(data As Date) As String
    Dim sql As String

    NomTaulaMovi = "V_Moviments_" & Format(data, "yyyy-mm")
   
    If Not ExisteixTaula(NomTaulaMovi) Then
        sql = "CREATE TABLE [" & NomTaulaMovi & "] ("
        sql = sql & "[Botiga] [float] NULL , "
        sql = sql & "[Data] [datetime] NULL , "
        sql = sql & "[Dependenta] [float] NULL ,"
        sql = sql & "[Tipus_moviment] [nvarchar] (25) NULL ,"
        sql = sql & "[Import] [float] NULL ,"
        sql = sql & "[Motiu] [nvarchar] (250) NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
    End If
   
End Function

Function NomTaulaAlertas() As String
    Dim sql As String

    NomTaulaAlertas = "Alertas"
   
    If Not ExisteixTaula(NomTaulaAlertas) Then
        sql = "CREATE TABLE [" & NomTaulaAlertas & "] ("
        sql = sql & "[Id] [nvarchar] (255) NOT NULL default (newid()), "
        sql = sql & "[Fecha] [datetime] NOT NULL default (getdate()), "
        sql = sql & "[Tipo] [nvarchar] (50) NOT NULL, "
        sql = sql & "[Texto] [nvarchar](255) NOT NULL, "
        sql = sql & "[Revisada] [int] NOT NULL default (0), "
        sql = sql & "[Param1] [nvarchar](255) NULL, "
        sql = sql & "[Param2] [nvarchar](255) NULL, "
        sql = sql & "[Param3] [nvarchar](255) NULL, "
        sql = sql & "[Param4] [nvarchar](255) NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
    End If
   
End Function

Function NomTaulaVentasBak(data As Date, empresaBD As String) As String
   
    NomTaulaVentasBak = "[V_Venut_" & Format(data, "yyyy-mm") & "]"
    If Not ExisteixTaula(NomTaulaVentasBak) Then
        NomTaulaVentasBak = empresaBD & "_Bak.dbo.[V_Venut_" & Format(data, "yyyy-mm") & "]"
    End If

End Function

Function TaulaTiquetMig(data As Date) As String
   Dim sql As String
   
   TaulaTiquetMig = "V_TiquetMig_" & Format(data, "yyyy-mm")
   
   If Not ExisteixTaula(TaulaTiquetMig) Then
        sql = "CREATE TABLE [" & TaulaTiquetMig & "] ("
        sql = sql & "[Botiga]          [float]    NULL, "
        sql = sql & "[Data]            [datetime] NULL, "
        sql = sql & "[Hora]            [int] NULL, "
        sql = sql & "[Vendes]          [float]    NULL, "
        sql = sql & "[Clients]         [float]    NULL, "
        sql = sql & "[TiquetMig]       [float]    NULL, "
        sql = sql & "[Tmst]            [datetime] NULL  "
        sql = sql & ") ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
   End If
   
End Function

Function NomTaulaVentas(data As Date) As String
   
    NomTaulaVentas = "V_Venut_" & Format(data, "yyyy-mm")
   
End Function

Function NomTaulaVentasPromo(data As Date) As String
    Dim sql As String
   
    NomTaulaVentasPromo = "V_Venut_Promo_" & Format(data, "yyyy-mm")
    If Not ExisteixTaula(NomTaulaVentasPromo) Then

        sql = "CREATE TABLE [" & NomTaulaVentasPromo & "]( "
        sql = sql & "[Botiga] [float] NULL, "
        sql = sql & "[Data] [datetime] NULL, "
        sql = sql & "[Dependenta] [float] NULL, "
        sql = sql & "[Num_tick] [float] NULL, "
        sql = sql & "[Estat] [nvarchar](25) NULL, "
        sql = sql & "[Plu] [float] NULL, "
        sql = sql & "[Quantitat] [float] NULL, "
        sql = sql & "[Import] [float] NULL, "
        sql = sql & "[Tipus_venta] [nvarchar](25) NULL, "
        sql = sql & "[FormaMarcar] [nvarchar](255) default ('') NULL, "
        sql = sql & "[Otros] [nvarchar](255) default ('') NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql

   End If
   
End Function
Function NomTaulaRevisats(data As Date) As String
   
   Dim sql As String
   
   NomTaulaRevisats = "V_Revisat_" & Format(data, "yyyy-mm")
   
   If Not ExisteixTaula(NomTaulaRevisats) Then
        sql = "CREATE TABLE [" & NomTaulaRevisats & "] ("
        sql = sql & "[Botiga]          [float]    NULL, "
        sql = sql & "[DataRevisio]     [datetime] NULL, "
        sql = sql & "[DataComanda]     [datetime] NULL, "
        sql = sql & "[Article]         [float]    NULL, "
        sql = sql & "[Viatge]          [nvarchar](255) NULL, "
        sql = sql & "[Equip]           [nvarchar](255) NULL, "
        sql = sql & "[Dependenta]      [float]    NULL, "
        sql = sql & "[Estat]           [nvarchar](25) NULL, "
        sql = sql & "[Aux]             [nvarchar](255) NULL "
        sql = sql & ") ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
   End If
   
   
End Function

Function NomTaulaFamiliesExtes() As String
   
   Dim sql As String
   
   NomTaulaFamiliesExtes = "FamiliesExtes"
   
   If Not ExisteixTaula(NomTaulaFamiliesExtes) Then
        sql = "CREATE TABLE [" & NomTaulaFamiliesExtes & "] ("
        sql = sql & "[Familia]  [nvarchar](255) NULL, "
        sql = sql & "[Variable] [nvarchar](255) NULL, "
        sql = sql & "[Valor]    [nvarchar](255) NULL "
        sql = sql & ") ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
   End If
   
   
End Function


Function NomTaulaVentasAoB(data As Date, AoB As Boolean) As String

    If AoB Then
        NomTaulaVentasAoB = NomTaulaVentasPrevistes(data)
    Else
        NomTaulaVentasAoB = NomTaulaVentas(data)
    End If
       
End Function

Function NomTaulaObjectius(data As Date) As String
   Dim sql As String
    
   NomTaulaObjectius = "V_Objectius_" & Format(data, "yyyy-mm")
   
   If Not ExisteixTaula(NomTaulaObjectius) Then
        sql = "CREATE TABLE [" & NomTaulaObjectius & "] ( "
        sql = sql & "[Botiga] [float] NULL, "
        sql = sql & "[Data] [datetime] NULL, "
        sql = sql & "[Periode] [nvarchar] (100) NULL, "
        sql = sql & "[Objectiu] [float] NULL, "
        sql = sql & "[Tipo] [nvarchar] (100) NULL "
        sql = sql & ") ON [PRIMARY]"
   
        ExecutaComandaSql sql
   End If
   
   
End Function

Function NomTaulaNumFacturaVendes(data As Date) As String
   Dim sql As String
    
   NomTaulaNumFacturaVendes = "NumFacturaVendes_" & Format(data, "yyyy")
   
   If Not ExisteixTaula(NomTaulaNumFacturaVendes) Then
        sql = "CREATE TABLE [" & NomTaulaNumFacturaVendes & "] ( "
        sql = sql & "[Data]    [datetime]   NULL, "
        sql = sql & "[Botiga] [float] NULL, "
        sql = sql & "[Factura] [float] NULL "
        sql = sql & ") ON [PRIMARY]"
   
        ExecutaComandaSql sql
   End If
   
End Function

Function NomTaulaRecepcioMP(data As Date) As String
   
   NomTaulaRecepcioMP = "V_RecepcioMateriesPrimeres_" & Format(data, "yyyy-mm")
   
End Function

Function DonamNomTaulaRecursos() As String
   Dim TaulaServit As rdoResultset, sql As String

   DonamNomTaulaRecursos = "Recursos"
   
   If Not ExisteixTaula(DonamNomTaulaRecursos) Then
      sql = "CREATE TABLE [" & DonamNomTaulaRecursos & "] ("
      sql = sql & "[Id] [nvarchar](255) NULL DEFAULT (newid()),"
      sql = sql & "[Nombre] [nvarchar](255) NULL,"
      sql = sql & "[Tipo] [nvarchar](255) NULL "
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function


Function DonamNomTaulaIncidencias() As String
   Dim TaulaServit As rdoResultset, sql As String
   
   DonamNomTaulaIncidencias = "Incidencias"
   
   If Not ExisteixTaula(DonamNomTaulaIncidencias) Then
      sql = "CREATE TABLE [" & DonamNomTaulaIncidencias & "] ("
      sql = sql & "[Id] [numeric](18, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,"
      sql = sql & "[TimeStamp] [datetime] NULL DEFAULT (getdate()),"
      sql = sql & "[Tipo] [nvarchar](100) NULL,"
      sql = sql & "[Usuario] [int] NULL,"
      sql = sql & "[Cliente] [nvarchar](255) NULL,"
      sql = sql & "[Recurso] [nvarchar](255) NULL,"
      sql = sql & "[Incidencia] [nvarchar](2500) NULL,"
      sql = sql & "[Estado] [nvarchar](50) NULL,"
      sql = sql & "[Observaciones] [nvarchar](2500) NULL,"
      sql = sql & "[FIniReparacion] [datetime] NULL,"
      sql = sql & "[FFinReparacion] [datetime] NULL,"
      sql = sql & "[Prioridad] [numeric](2, 0) NULL,"
      sql = sql & "[Tecnico] [int] NULL,"
      sql = sql & "[contacto] [nvarchar](250) NULL,"
      sql = sql & "[FProgramada] [datetime] NULL "
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function




Function DonamNomTaulaComandesCalculades(dia) As String
   Dim TaulaServit As rdoResultset, sql As String
   
   DonamNomTaulaComandesCalculades = "C_ComandaCalculada_" & Year(dia) & "_" & Right("00" & Month(dia), 2)
   
   If Not ExisteixTaula(DonamNomTaulaComandesCalculades) Then
      sql = "CREATE TABLE [" & DonamNomTaulaComandesCalculades & "] ("
      sql = sql & "[Id] [numeric](18, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,"
      sql = sql & "[Modificat]         [datetime] Default (GetDate()), "
      sql = sql & "[Client]            [float]          Null ,"
      sql = sql & "[Article]           [float]          Null ,"
      sql = sql & "[Viatge]            [nvarchar](255) NULL,"
      sql = sql & "[Equip]             [nvarchar](255) NULL,"
      sql = sql & "[Quantitat]         [float]          Null "
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function

Function DonamNomTaulaDedos() As String
   Dim TaulaServit As rdoResultset, sql As String
   
   DonamNomTaulaDedos = "Dedos"
   
   If Not ExisteixTaula(DonamNomTaulaDedos) Then
      sql = "CREATE TABLE [" & DonamNomTaulaDedos & "] ("
      sql = sql & "[usuario] [nvarchar](255) NULL,"
      sql = sql & "[fir] [nvarchar](4000) NULL,"
      sql = sql & "[userId] [numeric](18, 0) NULL"
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function

Function DonamNomTaulaTpvEquivalents() As String
   Dim TaulaServit As rdoResultset, sql As String
   DonamNomTaulaTpvEquivalents = "TpvEquivalents"
   
   If Not ExisteixTaula(DonamNomTaulaTpvEquivalents) Then
      sql = "CREATE TABLE [" & DonamNomTaulaTpvEquivalents & "] ("
      sql = sql & "[Tipus]       [nvarchar] (255) Null ,"
      sql = sql & "[valor1]      [nvarchar] (255) Null ,"
      sql = sql & "[valor2]      [nvarchar] (255) Null ,"
      sql = sql & "[valor3]      [nvarchar] (255) Null ,"
      sql = sql & "[valor4]      [nvarchar] (255) Null ,"
      sql = sql & "[valor5]      [nvarchar] (255) Null "
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function


Function DonamNomTaulaExportacio() As String
   Dim TaulaServit As rdoResultset, sql As String

   DonamNomTaulaExportacio = "Exportacions"
   
   If Not ExisteixTaula(DonamNomTaulaExportacio) Then
      sql = "CREATE TABLE [" & DonamNomTaulaExportacio & "] ("
      sql = sql & "[Tipus]      [nvarchar] (255) Null ,"
      sql = sql & "[Param1]      [nvarchar] (255) Null ,"
      sql = sql & "[Param2]      [nvarchar] (255) Null ,"
      sql = sql & "[Param3]      [nvarchar] (255) Null ,"
      sql = sql & "[Param4]      [nvarchar] (255) Null ,"
      sql = sql & "[Param5]      [nvarchar] (255) Null "
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql (sql)
   End If

End Function

Function DonamNomTaulaArchivoLines() As String
   Dim sql As String
   
   DonamNomTaulaArchivoLines = "ArchivoLines"
   
   If Not ExisteixTaula("ArchivoLines") Then
      sql = "CREATE TABLE [ArchivoLines] ( "
      sql = sql & "[id] [nvarchar] (255) NULL ,"
      sql = sql & "[nombre] [nvarchar] (20) NULL ,"
      sql = sql & "[fecha] [datetime] NULL ,"
      sql = sql & "[Linea] [nvarchar] (3000) NULL ,"
      sql = sql & "[NumLinea] [numeric](18, 0) NULL"
      sql = sql & ") ON [PRIMARY]"
      ExecutaComandaSql sql
   End If

End Function


Function DonamNomTaulaContabilitat() As String
   Dim sql As String
   
   DonamNomTaulaContabilitat = "ContaJb"
   
   If Not ExisteixTaula("ContaJb") Then
        sql = "CREATE TABLE [ContaJb] ("
        sql = sql & "[Id]            [nvarchar] (255) NULL ,"
        sql = sql & "[Origen]        [nvarchar] (255) NULL ,"
        sql = sql & "[Estat]         [nvarchar] (255) NULL ,"
        sql = sql & "[Data]          [nvarchar] (255) NULL ,"
        sql = sql & "[Asiento]       [nvarchar] (255) NULL ,"
        sql = sql & "[Fecha]         [nvarchar] (255) NULL ,"
        sql = sql & "[SubCuenta]     [nvarchar] (255) NULL ,"
        sql = sql & "[Contrapartida] [nvarchar] (255) NULL ,"
        sql = sql & "[Importe]       [nvarchar] (255) NULL ,"
        sql = sql & "[DebeHaber]     [nvarchar] (255) NULL ,"
        sql = sql & "[Factura]       [nvarchar] (255) NULL ,"
        sql = sql & "[Base]          [nvarchar] (255) NULL ,"
        sql = sql & "[Iva]           [nvarchar] (255) NULL ,"
        sql = sql & "[RecEquiv]      [nvarchar] (255) NULL ,"
        sql = sql & "[Documento]     [nvarchar] (255) NULL ,"
        sql = sql & "[Departamento]  [nvarchar] (255) NULL ,"
        sql = sql & "[Clave]         [nvarchar] (255) NULL ,"
        sql = sql & "[Eestado]       [nvarchar] (255) NULL ,"
        sql = sql & "[Ncasado]       [nvarchar] (255) NULL ,"
        sql = sql & "[Tcasado]       [nvarchar] (255) NULL ,"
        sql = sql & "[Trans]         [nvarchar] (255) NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
   End If

End Function
