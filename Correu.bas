Attribute VB_Name = "Correu"
Option Explicit


Global Q_FileExisteix  As rdoQuery, Q_FileInserta As rdoQuery, Q_File_Carregada As rdoQuery


Dim connMDB 'Cadena conexion Access
Dim ultimaSql As String
'**********************************************************************************
'************************************************************
'* Funciones necesarias
'************************************************************
Sub AreglaNomsMalPosats()
    Dim Fil As String, f, L As String, NomVell() As String, NomNou() As String, i As Integer, Contingut As String, NomMaquina As String, nomfile As String, K
    Dim NomTaula As String
    Dim kr As Integer
    ReDim NomVell(0)
    ReDim NomNou(0)
Exit Sub
    f = FreeFile
    Fil = Dir(AppPath & "\*.SqlTrans")
    kr = 0
    While Len(Fil) > 0
        kr = kr + 1
        DescomposaContingut Fil, Contingut, NomMaquina, nomfile
        Open AppPath & "\" & Fil For Input As #f
        K = 0
        While Not EOF(f) And K < 10
            Line Input #f, L
            If Left(L, 14) = "[Sql-NomTaula:" Then
                NomTaula = DonamParam(L)
                If InStr(Fil, NomTaula) = 0 Then
                
                ReDim Preserve NomVell(UBound(NomVell) + 1)
                ReDim Preserve NomNou(UBound(NomNou) + 1)
                NomVell(UBound(NomVell)) = Fil
                NomNou(UBound(NomNou)) = ""
                
                End If
                
                K = 11
            End If
        Wend
        Close f
        Fil = Dir()
    Wend

    For i = 1 To UBound(NomVell)
        Name UBound(NomVell) As UBound(NomVell)
    Next
   

End Sub

Function BuscaLastData() As Date
   Dim Rs As rdoResultset
      
   If Not ExisteixTaula("RecordsFiles") Then ExecutaComandaSql "CREATE TABLE [RecordsFiles] ([Path] [nvarchar] (255) NULL ,[Nom] [nvarchar] (255) NULL ,[Tot] [nvarchar] (255) NULL ,[Data] [datetime] NULL,[DataRebut] [datetime] NULL ,[DataEnviat] [datetime] NULL,[Kb] [float] NULL ) "
   If Not ExisteixTaula("RecordsFilesBak") Then ExecutaComandaSql "CREATE TABLE [RecordsFilesBak] ([Path] [nvarchar] (255) NULL ,[Nom] [nvarchar] (255) NULL ,[Tot] [nvarchar] (255) NULL ,[Data] [datetime] NULL,[DataRebut] [datetime] NULL ,[DataEnviat] [datetime] NULL,[Kb] [float] NULL ) "
   ExecutaComandaSql "CREATE   INDEX [RebordsBackIndex] ON [dbo].[RecordsFiles] ([Tot], [Kb]) ON [PRIMARY]"
   Set Q_FileExisteix = Db.CreateQuery("", "Select DataRebut,DataEnviat From RecordsFiles where Tot = ? And Kb = ?")
   Set Q_FileInserta = Db.CreateQuery("", "Insert Into RecordsFiles ([Path],[Nom],[Tot],[Data],[DataRebut],[DataEnviat],[Kb]) Values (?,?,?,?,Null,Null,?) ")
   Set Q_File_Carregada = Db.CreateQuery("", "Update RecordsFiles Set DataRebut = GetDate() Where Tot = ? And Kb = ?")
   
   BuscaLastData = DateAdd("d", -15, Now)
   Set Rs = Db.OpenResultset("Select * From Records Where Concepte = 'CorreoRebutFins'")
   If Not Rs.EOF Then If Not IsNull(Rs(0)) Then BuscaLastData = Rs(0)
   Rs.Close

End Function

Sub CreaEmpressa(nom As String, Nomllarg As String)
   Dim Path As String
   Dim Q As rdoQuery
   
'CREATE PROCEDURE CreaEmpresa
'    @NomDb nvarchar(255)
'AS
'begin
'EXEC ("CREATE DATABASE [" + @NomDb + "]  ON (NAME = N'" + @NomDb + "_dat', FILENAME = N'E:\Data\" + @NomDb + ".mdf' , SIZE = 1, FILEGROWTH = 10%) LOG ON (NAME = N'" + @NomDb + "_log', FILENAME = N'E:\Data\" + @NomDb + ".ldf' , SIZE = 2, FILEGROWTH = 10%)" )
'End
'GO
   
   
'Crea Db
    Path = "E:\Data\"
    ExecutaComandaSql "Use Hit"
    ExecutaComandaSql "Execute hit.dbo.CreaEmpresa '" & nom & "'"
    ExecutaComandaSql "Use Master"
    
'    Set Q = Db.CreateQuery("", "CREATE DATABASE [" & Nom & "]  ON (NAME = N'" & Nom & "_dat', FILENAME = N'" & Path & "" & Nom & ".mdf' , SIZE = 683, FILEGROWTH = 10%) LOG ON (NAME = N'" & Nom & "_log', FILENAME = N'" & Path & "" & Nom & ".ldf' , SIZE = 2, FILEGROWTH = 10%)")
'    Q.Execute dbRunAsync
'
'    While Q.StillExecuting
'       DoEvents
'    Wend
'    ExecutaComandaSql "CREATE DATABASE [" & Nom & "]  ON (NAME = N'" & Nom & "_dat', FILENAME = N'" & Path & "" & Nom & ".mdf' , SIZE = 683, FILEGROWTH = 10%) LOG ON (NAME = N'" & Nom & "_log', FILENAME = N'" & Path & "" & Nom & ".ldf' , SIZE = 2, FILEGROWTH = 10%)"
    
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'autoclose', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'bulkcopy', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'trunc. log', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'torn page detection', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'read only', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'dbo use', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'single', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'autoshrink', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'ANSI null default', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'recursive triggers', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'ANSI nulls', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'concat null yields null', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'cursor close on commit', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'default to local cursor', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'quoted identifier', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'ANSI warnings', N'false'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'auto create statistics', N'true'"
    ExecutaComandaSql "exec sp_dboption N'" & nom & "', N'auto update statistics', N'true'"
    ExecutaComandaSql "if( ( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) ) or ( (@@microsoftversion / power(2, 24) = 7) and (@@microsoftversion & 0xffff >= 1082) ) )  exec sp_dboption N'" & nom & "', N'db chaining', N'false' "

' Crea access
   ExecutaComandaSql "Use Hit"
   ExecutaComandaSql "insert into web_empreses (Nom,Descripcio,Db,Logo,Estil,Path,Activa,Llicencia,Servidor,Ppv,publicitat) values ('" & nom & "','" & Nomllarg & "','" & nom & "','" & nom & "','demo.css','\Empreses\' + '" & nom & "','Si',0,'" & ServerActual & "','[comandes]',1) "
   ExecutaComandaSql "insert into web_serveiscomuns (Tipus,Empresa,Actiu) values (1,'" & nom & "',0)"
   ExecutaComandaSql "insert into web_serveiscomuns (Tipus,Empresa,Actiu) values (3,'" & nom & "',0)"
   ExecutaComandaSql "insert into web_users (Nom,Empresa,Password,NivellSeguretat,PaginaInici,Route_ip,Route_Cookie,Route_Url) values ('" & nom & "','" & nom & "','Inicial',9,'/resultados/','','','')"
   
'Crea Taules
    ExecutaComandaSql "Use " & nom
    
    ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL ,[TimeStamp] [datetime] NULL ,  [TaulaOrigen] [nvarchar] (255) NULL ) ON [PRIMARY] "
    ExecutaComandaSql "CREATE TABLE [Articles] ([Codi] [decimal](18, 0) NULL ,  [NOM] [nvarchar] (255) NULL ,   [PREU] [float] NULL ,   [PreuMajor] [float] NULL ,  [Desconte] [float] NULL ,   [EsSumable] [bit] NULL ,    [Familia] [nvarchar] (255) NULL ,   [CodiGenetic] [int] NULL ,  [TipoIva] [float] NULL ,    [NoDescontesEspecials] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Articles_Zombis] ( [TimeStamp] [datetime] NULL ,   [Codi] [numeric](18, 0) NULL ,  [NOM] [nvarchar] (255)  NULL ,  [PREU] [float] NOT NULL ,   [PreuMajor] [float] NULL ,  [Desconte] [float] NULL ,   [EsSumable] [bit] NOT NULL ,    [Familia] [nvarchar] (255)  NULL ,  [CodiGenetic] [int] NOT NULL ,  [TipoIva] [float] NULL ,    [NoDescontesEspecials] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Atributs] (    [Codi] [float] NOT NULL ,   [Nom] [nvarchar] (255) NULL ,   [TexteAnexat] [nvarchar] (255)  NULL ,  [Es Prefabricat] [bit] NOT NULL ,   [ModificaPreu1] [bit] NOT NULL ,    [Increment_Preu_1] [nvarchar] (255)  NULL , [Increment_Pct_1] [nvarchar] (255)  NULL ,  [ModificaPreu2] [bit] NOT NULL ,    [Increment_Preu_2] [nvarchar] (255) NULL ,  [Increment_Pct_2] [nvarchar] (255)  NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Clients] ( [Codi] [int] NULL , [Nom] [nvarchar] (255)  NULL ,  [Nif] [nvarchar] (255)  NULL ,  [Adresa] [nvarchar] (255)  NULL ,   [Ciutat] [nvarchar] (255)  NULL ,   [Cp] [nvarchar] (255)  NULL ,   [Lliure] [nvarchar] (255)  NULL ,   [Nom Llarg] [nvarchar] (255)  NULL ,    [Tipus Iva] [int] NULL ,    [Preu Base] [int] NULL ,    [Desconte ProntoPago] [int] NULL ,  [Desconte 1] [int] NULL ,   [Desconte 2] [int] NULL ,   [Desconte 3] [int] NULL ,   [Desconte 4] [int] NULL ,   [Desconte 5] [int] NULL ,   [AlbaraValorat] [int] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ClientsFinals] (   [Id] [nvarchar] (255) NULL ,    [IdExterna] [nvarchar] (255) NULL , [Nom] [nvarchar] (255) NULL ,   [Nif] [nvarchar] (255) NULL ,   [Telefon] [nvarchar] (255) NULL ,   [Adreca] [nvarchar] (255) NULL ,    [emili] [nvarchar] (255) NULL , [Descompte] [nvarchar] (255) NULL , [Altres] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ComandesMemotecnicPerClient] ( [TimeStamp] [datetime] NULL ,   [CodiArticle] [numeric](18, 0) NULL ,   [Viatge] [nvarchar] (50) NULL , [Equip] [nvarchar] (50) NULL ,  [Client] [numeric](18, 0) NULL ,    [Prob] [numeric](18, 0) NULL ,  [Pct] [numeric](18, 0) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ComandesParams] (  [Tipus] [numeric](18, 0) NULL , [Camp] [numeric](18, 0) NULL ,  [Valor] [numeric](18, 0) NULL , [Descripcio] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ComandesPlantilles] (  [Nom] [nvarchar] (50) NULL ,    [Pos] [numeric](18, 0) NULL ,   [Article] [numeric](18, 0) NULL ,   [viatge] [nvarchar] (50) NULL , [Equip] [nvarchar] (50) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Compres] ( [Data] [datetime] NULL ,    [Proveidor] [nvarchar] (255) NULL , [Producte] [numeric](18, 0) NULL ,  [Quantitat] [float] NULL ,  [PreuCompra] [float] NULL , [PreuVenta] [float] NULL ,  [MargeEsperat] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ConstantsClient] ( [Codi] [numeric](18, 0) NULL ,  [Variable] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [Valor] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ConstantsEmpresa] (    [Camp] [nvarchar] (255) NULL ,  [Valor] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Dependentes] ( [CODI] [int] NULL , [NOM] [nvarchar] (255) NULL ,   [MEMO] [nvarchar] (255) NULL ,  [TELEFON] [nvarchar] (255) NULL ,   [ADREÇA] [nvarchar] (255) NULL ,    [Icona] [nvarchar] (255) NULL , [Hi Editem Horaris] [int] NULL ,    [Tid] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Dependentes_Hores_Preferencies] (  [Camp] [nvarchar] (255) NULL ,  [Valor] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Dependentes_Laboral] ( [Id] [uniqueidentifier] NULL ,  [TimeStamp] [datetime] NULL ,   [Codi] [int] NULL , [EquipTreball] [nvarchar] (255) NULL ,  [NHoresBase] [float] NULL , [PreuHoraBase] [float] NULL ,   [PreuHoraExtra] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Dependentes_Zombis] (  [TimeStamp] [datetime] NULL ,   [CODI] [int] NULL , [NOM] [nvarchar] (20) NULL ,    [MEMO] [nvarchar] (8) NULL ,    [TELEFON] [nvarchar] (10) NULL ,    [ADREÇA] [nvarchar] (30) NULL , [Icona] [nvarchar] (50) NULL ,  [Hi Editem Horaris] [int] NULL ,    [Tid] [nvarchar] (50) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [dependentesExtes] (    [id] [nvarchar] (255) NULL ,    [nom] [nvarchar] (255) NULL ,   [valor] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [DeutesAnticips] (  [Id] [nvarchar] (255) NULL ,    [Dependenta] [float] NULL , [Client] [nvarchar] (255) NULL ,    [Data] [datetime] NULL ,    [Estat] [float] NULL ,  [Tipus] [float] NULL ,  [Import] [float] NULL , [Botiga] [float] NULL , [Detall] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [EquipsDeTreball] ( [Nom] [nvarchar] (255) NULL ,   [Defecte] [bit] NOT NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Exportacions] (    [Tipus] [nvarchar] (255) NULL , [Param1] [nvarchar] (255) NULL ,    [Param2] [nvarchar] (255) NULL ,    [Param3] [nvarchar] (255) NULL ,    [Param4] [nvarchar] (255) NULL ,    [Param5] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [FacturacioComentaris] (    [IdFactura] [nvarchar] (255) NULL , [Data] [datetime] NULL ,    [Comentari] [nvarchar] (255) NULL , [Cobrat] [nvarchar] (1) NULL CONSTRAINT [DF_FacturacioComentaris_Cobrat] DEFAULT ('S')) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Families] (    [Nom] [nvarchar] (255) NULL ,   [Pare] [nvarchar] (255) NULL ,  [Estatus] [nvarchar] (255) NULL ,   [Nivell] [int] NOT NULL ,   [Utilitza] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [FeinesAFer] (  [Id] [nvarchar] (255) NULL CONSTRAINT [DF_FeinesAFer_Id] DEFAULT (newid()), [Tipus] [nvarchar] (255) NULL , [Ciclica] [nvarchar] (255) NULL ,   [Param1] [nvarchar] (255) NULL ,    [Param2] [nvarchar] (255) NULL ,    [Param3] [nvarchar] (255) NULL ,    [Param4] [nvarchar] (255) NULL ,    [Param5] [nvarchar] (255) NULL, tmStmp [datetime] NULL CONSTRAINT [DF_FeinesAFer_TmStmp] DEFAULT (getdate()) ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Festius] ( [Fecha] [varchar] (10) NOT NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Formules] (    [NomMasa] [nvarchar] (255) NULL ,   [Comentaris] [nvarchar] (1000) NULL ,   [Components] [nvarchar] (1000) NULL ,   [TimeStamp] [datetime] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [InteresaContingut] (   [Nom] [nvarchar] (255) NULL ,   [LaAgafem] [int] NULL , [LaEsborrem] [int] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [inventaris_comandes] ( [COD_PRODUCTO] [int] NOT NULL , [COD_BOTIGA] [int] NOT NULL ,   [DATA] [datetime] NULL ,    [COMANDA] [int] NULL ,  [PROVEEDOR] [nvarchar] (50) NULL ,  [PREU] [float] NULL ,   [MARGE] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Inventaris_Stocks_Minims] (    [tmst] [datetime] NOT NULL CONSTRAINT [DF_Inventaris_Stocks_Minims_tmst] DEFAULT (getdate()),   [article] [decimal](18, 0) NULL ,   [botiga] [int] NULL ,   [stock_min] [int] NULL ,    [proveedor] [nvarchar] (255) NULL , [precio] [float] NULL , [margen] [decimal](18, 0) NULL ,    [ultima_compra] [datetime] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [IvasAPagar] (  [Client] [float] NULL , [NumFactura] [float] NULL , [DataFactura] [datetime] NULL , [ClientNom] [varchar] (255) NULL ,  [ClientNif] [varchar] (255) NULL ,  [ClientAdresa] [varchar] (255) NULL ,   [ClientCiutat] [varchar] (255) NULL ,   [ClientCp] [varchar] (255) NULL ,   [ClientLliure] [varchar] (255) NULL ,   [Atribut] [int] NULL , " _
        & " [EmpresaNom] [varchar] (255) NULL , [EmpresaNif] [varchar] (255) NULL , [EmpresaAdresa] [varchar] (255) NULL ,  [EmpresaCiutat] [varchar] (255) NULL ,  [EmpresaCp] [varchar] (255) NULL ,  [EmpresaLliure] [varchar] (255) NULL ,  [Titol1] [varchar] (255) NULL , [Base1] [varchar] (255) NULL ,  [IvaPercent1] [varchar] (255) NULL ,    [Iva1] [varchar] (255) NULL ,   [RetencioPercent1]" _
        & "[varchar] (255) NULL ,   [Retencio1] [varchar] (255) NULL ,  [Titol2] [varchar] (255) NULL , [Base2] [varchar] (255) NULL ,  [IvaPercent2] [varchar] (255) NULL ,    [Iva2] [varchar] (255) NULL ,   [RetencioPercent2] [varchar] (255) NULL ,   [Retencio2] [varchar] (255) NULL , [Titol3] [varchar] (255) NULL , [Base3] [varchar] (255) NULL ,  [IvaPercent3] [varchar] (255) NULL ,    [Iva3] " _
        & "[varchar] (255) NULL ,   [RetencioPercent3] [varchar] (255) NULL ,   [Retencio3] [varchar] (255) NULL ,  [Titol4] [varchar] (255) NULL , [Base4] [varchar] (255) NULL ,  [IvaPercent4] [varchar] (255) NULL ,    [Iva4] [varchar] (255) NULL ,   [RetencioPercent4] [varchar] (255) NULL ,   [Retencio4] [varchar] (255) NULL ,  [Total] [varchar] (255) NULL ,  [BaseBrutaFactura] [varchar] (255) NULL" _
        & ",   [Impostos] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Masas] (   [Grup] [nvarchar] (255) NULL ,  [Article] [int] NULL ,  [Factor] [float] NULL , [Viatge] [nvarchar] (255) NULL ,    [Atribut] [int] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Memotecnics] ( [Tpv] [bit] NOT NULL ,  [Memotecnic] [nvarchar] (255) NULL ,    [Codi] [int] NOT NULL , [Viatge] [nvarchar] (255) NULL ,    [Equip] [nvarchar] (255) NULL , [Atribut] [int] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Missatges] (   [Id] [nvarchar] (100) NULL CONSTRAINT [DF_Missatges_Id] DEFAULT (newid()),  [TimeStamp] [datetime] NULL CONSTRAINT [DF_Missatges_TimeStamp] DEFAULT (getdate()),    [QuiStamp] [nvarchar] (255) NULL CONSTRAINT [DF_Missatges_QuiStamp] DEFAULT (N'Web'),   [DataEnviat] [datetime] NULL CONSTRAINT [DF_Missatges_DataEnviat] DEFAULT (getdate()),  [DataRebut] [datetime] NULL ,   [Desti] [nvarchar] (255) NULL , [Origen] [nvarchar] (255) NULL ,    [Texte] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [MissatgesAEnviar] ([Tipus] [varchar] (255) NULL ,  [Param] [varchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [OrdenReparto] (    [Codi] [real] NOT NULL ,    [Tipus] [nvarchar] (50) NOT NULL ,  [Reparto] [nvarchar] (50) NOT NULL ,    [Nom] [nvarchar] (50) NOT NULL ,    [Orden] [int] NOT NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Pagos] (   [Id] [uniqueidentifier] NULL ,  [TimeStamp] [datetime] NULL ,   [QuiStamp] [nvarchar] (255) NULL ,  [DataPagat] [datetime] NULL ,   [Client] [float] NULL , [Import] [float] NULL , [Texte] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ParamsComandesCfg] (   [Client] [float] NULL , [Tipus] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ParamsHw] (    [Tipus] [numeric](18, 0) NULL , [Codi] [numeric](18, 0) NULL ,  [Valor1] [nvarchar] (255) NULL ,    [Valor2] [nvarchar] (255) NULL ,    [Valor3] [nvarchar] (255) NULL ,    [Valor4] [nvarchar] (255) NULL ,    [Descripcio] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [ParamsTpv] (   [CodiClient] [int] NULL ,   [Variable] [nvarchar] (255) NULL ,  [Valor] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [PreferenciasOrden] (   [Equipo] [nvarchar] (50) NULL , [Masa] [nvarchar] (50) NULL ,   [Tipo] [nvarchar] (50) NULL ,   [Codigo] [int] NULL ,   [Orden] [int] NOT NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [promocions] (  [di] [datetime] NULL ,  [df] [datetime] NULL ,  [botiga] [numeric](18, 0) NULL ,    [article] [numeric](18, 0) NULL ,   [punts] [numeric](18, 0) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Punts] (   [IdClient] [nvarchar] (255) NULL ,  [Punts] [float] NULL ,  [data] [datetime] NULL ,    [Punts2] [float] NULL , [data2] [datetime] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [PuntsAcumulats] (  [Client] [nvarchar] (255) NULL ,    [Botiga] [float] NULL , [Dependenta] [float] NULL , [Num_tick] [float] NULL ,   [Data] [datetime] NULL ,    [Import] [float] NULL , [Punts] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [QueTinc] ( [QueEs] [nvarchar] (255) NULL , [QuinEs] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Records] ( [TimeStamp] [datetime] NULL ,   [Concepte] [nvarchar] (255) NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [RecordsFiles] (    [Path] [nvarchar] (255) NULL ,  [Nom] [nvarchar] (255) NULL ,   [Tot] [nvarchar] (255) NULL ,   [Data] [datetime] NULL ,    [DataRebut] [datetime] NULL ,   [DataEnviat] [datetime] NULL ,  [Kb] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [TarifesEspecials] (    [TarifaCodi] [int] NULL ,   [TarifaNom] [nvarchar] (20) NULL ,  [Codi] [int] NULL , [PREU] [float] NOT NULL ,   [PreuMajor] [float] NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [TipusIva] (    [Tipus] [nvarchar] (255) NULL , [Iva] [float] NOT NULL ,    [Irpf] [float] NOT NULL ) ON [PRIMARY]"
    ExecutaComandaSql "CREATE TABLE [Viatges] ( [Nom] [nvarchar] (255) NULL ,   [MinutInici] [int] NOT NULL ,   [Defecte] [bit] NOT NULL ) ON [PRIMARY]"
'Valors Inicials
    ExecutaComandaSql "Delete Articles where codi = 1"
    ExecutaComandaSql "insert into articles ([Codi], [NOM], [PREU], [PreuMajor], [Desconte], [EsSumable], [Familia], [CodiGenetic], [TipoIva], [NoDescontesEspecials] ) values (1,'Barra 1/2' ,1,1,1,1,'',1,1,1)"
    ExecutaComandaSql "insert into viatges (minutinici,defecte) values ('Inicial',120,1)"
    ExecutaComandaSql "insert into equipsdetreball (defecte) values ('Inicial',1)"
    ExecutaComandaSql "delete families"
    ExecutaComandaSql "insert into families (Nom,Pare,Estatus,Nivell,Utilitza) values ('Article','',0,0,'')"
    ExecutaComandaSql "insert into families (Nom,Pare,Estatus,Nivell,Utilitza) values ('Pa','Article',4,1,'')"
    ExecutaComandaSql "insert into families (Nom,Pare,Estatus,Nivell,Utilitza) values ('Bolleria','Article',4,1,'')"
    ExecutaComandaSql "insert into families (Nom,Pare,Estatus,Nivell,Utilitza) values ('Blanc','Pa',4,2,'')"
    ExecutaComandaSql "insert into families (Nom,Pare,Estatus,Nivell,Utilitza) values ('Barres','Blanc',4,3,'')"

    ExecutaComandaSql "delete tipusiva"
    ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (1,4,0.5) "
    ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (2,7,1) "
    ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (3,16,4) "
    'ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (2,8,1) "
    'ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (3,18,4) "
    ExecutaComandaSql "insert into tipusiva (tipus,iva,irpf) values (4,0,0) "
    ExecutaComandaSql "Use Hit"
  
End Sub


Sub EnviaComandesBotiga(Files() As String, botiga As String)
   Dim D As Date, Rs As rdoResultset, Botigues() As Double, i As Integer
   
   ReDim Files(0)
   
   InformaMiss "Preparant Comandes " & BotigaCodiNom(botiga), False
   
   EnviaComandes Files, botiga
   
End Sub



Function EnviaFacturacioCreaFile(Files, f, CampsClaus, Final As String) As String
   Dim Coletilla As String, nomfile As String
   
On Error Resume Next
   Close f
On Error GoTo 0

   Coletilla = ""
   Do
      nomfile = AppPath & "\" & Format(Now, "yyyy-mm-ddhhmmss") & Coletilla & Final
      Coletilla = Val(Coletilla) + 1
   Loop While Len(Dir(nomfile)) > 0
   
   Open nomfile For Output As #f
   
   If Len(CampsClaus) > 0 Then Print #f, "[Sql-LlistaCamps:" & CampsClaus & "]"
   
   ReDim Preserve Files(UBound(Files) + 1)
   Files(UBound(Files)) = nomfile
   
   EnviaFacturacioCreaFile = nomfile

End Function

Function FixaNumLlicencia(n As String) As Boolean

    frmSplash.IpConexio.Cfg_Llicencia = n
   
End Function


Function DescomposaContingut(Item As String, Contingut As String, NomMaquina As String, nomfile As String)
   Dim P As Integer, p1 As Integer
   
   Contingut = ""
   P = InStr(Item, "[Contingut#")
   If P > 0 Then
      P = P + 11
      p1 = InStr(P, Item, "]")
      Contingut = Mid(Item, P, p1 - P)
   End If
   
   NomMaquina = ""
   P = InStr(Item, "[Maquina#")
'   If P = 0 Then   ' Concordia
'        Item = "[Maquina#00233]" & Item
'        P = InStr(Item, "[Maquina#")
'   End If
   If P > 0 Then
      P = P + 9
      p1 = InStr(P, Item, "]")
      NomMaquina = Mid(Item, P, p1 - P)
   End If
   
   
   nomfile = Right(Item, Len(Item) - (P + Len(NomMaquina)))
   
End Function

Sub FtpCopy()
   Dim Rs As rdoResultset, Server, User, Psw, Fil, Fil1, Fil2, Fil3
   Dim IpFtpExtern As FTP
   
On Error GoTo res

   Set Rs = Db.OpenResultset("Select * from feinesafer where tipus='FtpCopy' ")
   
   If Rs.EOF Then Exit Sub
   If IsNull(Rs("Param1")) Then Exit Sub
   Server = Rs("Param1")
   If IsNull(Rs("Param2")) Then Exit Sub
   User = Rs("Param2")
   If IsNull(Rs("Param3")) Then Exit Sub
   Psw = Rs("Param3")
   
'FTP.casaametller.net
'usuari: sdx1188
'password: Ametller6259

   If Server = "" Or User = "" Or Psw = "" Then Exit Sub
   
   Fil1 = Dir(AppPath & "\*.SqlTrans")
   Fil2 = Dir(AppPath & "\Msg\Cfg\" & "\*.SqlTrans")
   Fil3 = Dir(AppPath & "\Msg\" & "\*.SqlTrans")
   If Fil1 = "" And Fil2 = "" And Fil3 = "" Then Exit Sub
   
   Set IpFtpExtern = frmSplash.FTP1
   
   IpFtpExtern.WinsockLoaded = True
   IpFtpExtern.RemoteHost = Server
   IpFtpExtern.User = User
   IpFtpExtern.Password = Psw
   IpFtpExtern.Passive = True
   Informa "Intentant Logon a ." & Server
   IpFtpExtern.Action = a_Logon
   Informa "Logon Ok. " & Server
   
   Fil = Dir(AppPath & "\*.SqlTrans")
   While Len(Fil) > 0
        IpFtpExtern.RemoteFile = Fil
        IpFtpExtern.LocalFile = AppPath & "\" & Fil
        IpFtpExtern.TransferMode = tm_Binary
        Informa "Pujant " & Fil
        IpFtpExtern.Action = a_Upload
        Fil = Dir
   Wend
   
   Fil = Dir(AppPath & "\Msg\Cfg\*.SqlTrans")
   While Len(Fil) > 0
        IpFtpExtern.RemoteFile = Fil
        IpFtpExtern.LocalFile = AppPath & "\Msg\Cfg\" & Fil
        IpFtpExtern.TransferMode = tm_Binary
        Informa "Pujant " & Fil
        IpFtpExtern.Action = a_Upload
        Fil = Dir
   Wend
   
   Fil = Dir(AppPath & "\Msg\*.SqlTrans")
   While Len(Fil) > 0
        IpFtpExtern.RemoteFile = Fil
        IpFtpExtern.LocalFile = AppPath & "\Msg\" & Fil
        IpFtpExtern.TransferMode = tm_Binary
        Informa "Pujant " & Fil
        IpFtpExtern.Action = a_Upload
        Fil = Dir
   Wend
   
   IpFtpExtern.Action = a_Logoff
res:
End Sub

Sub GeneraClients(Tipus() As String, Params() As String, Files() As String, Optional Tots As Boolean = False)
    Dim sql As String, i As Integer, Rs As rdoResultset, Fils() As String
    
    ExecutaComandaSql "Drop table ClientFinalsPendentsEnviar"
    ExecutaComandaSql "CREATE TABLE [ClientFinalsPendentsEnviar] ([Cli] [nvarchar] (255) NULL)"
    ExecutaComandaSql "Delete ClientsFinalsPropietats Where Id = ''  "
   
    For i = 0 To UBound(Tipus)
        If Tipus(i) = "ClientsFinals" Then ExecutaComandaSql "Insert into [ClientFinalsPendentsEnviar] (Cli) Values ('" & Params(i) & "')"
    Next
   

    If Tots Then
        sql = "Select Id, IdExterna, Nom, Nif, Telefon, Adreca, emili, Descompte, Altres From ClientsFinals "
    Else
        sql = "Select Id, IdExterna, Nom, Nif, Telefon, Adreca, emili, Descompte, Altres From ClientsFinals Where Id in (Select Cli From ClientFinalsPendentsEnviar)"
    End If
   
    GemeraSqlTrans "ClientsFinals", Files, sql
    
    If Tots Then
        sql = "Select Id, Variable, Valor From ClientsFinalsPropietats "
    Else
        sql = "Select Id, Variable, Valor From ClientsFinalsPropietats Where Id in (Select Cli From ClientFinalsPendentsEnviar)"
    End If
    
    ReDim Fils(0)
            
    GemeraSqlTrans "ClientsFinalsPropietats", Fils, sql
    If UBound(Fils) > 0 Then
       ReDim Preserve Files(UBound(Files) + 1)
       Files(UBound(Files)) = Fils(1)
    End If
    
End Sub
Sub GemeraSqlAccio(Param As String, Files() As String, Contingut As String)
   Dim Dada As String, f, codiBotiga As Double
   Dim P As Integer, tros As String
   
   codiBotiga = -1
   P = InStr(Param, ":")
   If P > 0 Then codiBotiga = Left(Param, P - 1)
   Param = Right(Param, Len(Param) - P)
   P = InStr(Param, "-")
   If P > 0 Then Mid(Param, P, 1) = ":"
   P = InStr(Param, "-")
   If P > 0 Then Mid(Param, P, 1) = ":"
   Dada = Param
   
   ReDim Files(1)
   Files(1) = AppPath & "\Accions_Botiga_" & codiBotiga & "_Accions.SqlTrans"
   On Error Resume Next
      FitcherProcesat Files(1), True, True
   On Error GoTo 0
   f = FreeFile
   Open Files(1) For Output As #f
   Print #f, "[CalEnviar:ElDia_" & Dada & "]"
   Close f
      
   Contingut = "Accions_Botiga_" & codiBotiga
   
End Sub
Sub GemeraSqlAccioTots(Tip() As String, Param() As String, Files() As String, Contingut As String)
   Dim Dada As String, f, codiBotiga As Double, i As Integer
   Dim P As Integer, tros As String
   
   ReDim Files(1)
   Files(1) = AppPath & "\Accions_TotesLesBotigues.SqlTrans"
   On Error Resume Next
      FitcherProcesat Files(1), True, True
   On Error GoTo 0
   f = FreeFile
   Open Files(1) For Output As #f
   For i = 0 To UBound(Param)
      If Tip(i) = "ClientFinal_Esborrat" Or Tip(i) = "Deute_Cambiat" Then Print #f, "[" & Tip(i) & ":" & Param(i) & "]"
   Next
   
   Close f
   Contingut = "Accions_TotesLesBotigues"
   
End Sub

Function Idt_CodiDependenta(IdT As String) As Double
   Idt_CodiDependenta = -1
   
   
   
   
End Function

Function LlicenciaCodiClient(Llic As String) As Double
   Dim Rs As rdoResultset
   
   LlicenciaCodiClient = -1
   If LlicenciaCodiClient = -1 And Not Llic = "" Then
      Set Rs = Db.OpenResultset("Select Valor1 From ParamsHw Where Codi = '" & Llic & "'  And Tipus = 1 ")
      If Not Rs.EOF Then
        If Not IsNull(Rs("Valor1")) Then
            LlicenciaCodiClient = Rs("Valor1")
         End If
      End If
      Rs.Close
   End If
   
   If LlicenciaCodiClient = -1 And Not Llic = "" Then
      Set Rs = Db.OpenResultset("Select CodiClient From ParamsTpv Where Variable = 'NomMaquina' And Valor = '" & Llic & "' ")
      If Not Rs.EOF Then
         If Not IsNull(Rs("CodiClient")) Then
            LlicenciaCodiClient = Rs("CodiClient")
         Else
            LlicenciaCodiClient = Llic
         End If
         ExecutaComandaSql "Insert into Clients (codi,nom,nif,adresa ,[Desconte 5]) values (" & LlicenciaCodiClient & ",'Botiga Llicencia " & LlicenciaCodiClient & "','','',0)"
      End If
      Rs.Close
   End If
   

End Function
Function CodiClientLlicencia(Llic As String) As Double
   Dim Rs As rdoResultset
   
   CodiClientLlicencia = Llic
   Set Rs = Db.OpenResultset("Select Codi From ParamsHw Where Valor1 = '" & Llic & "'  And Tipus = 1 ")
   If Not Rs.EOF Then
     If Not IsNull(Rs("Codi")) Then
         CodiClientLlicencia = Rs("Codi")
      End If
   End If
   Rs.Close

End Function

Function CodiClientTeTpv(Codi As Double) As Boolean
   Dim Rs As rdoResultset
   
   CodiClientTeTpv = False
      
   Set Rs = Db.OpenResultset("Select Codi From ParamsHw Where Valor1 = '" & Codi & "'  And Tipus = 1 ")
   If Not Rs.EOF Then If Not IsNull(Rs("Codi")) Then If Rs("codi") > 0 Then CodiClientTeTpv = True
   Rs.Close

End Function

Function LlicenciaValida(Lic As String, NumTel As String, User As String, Psw As String, err As String) As Boolean
   
   If frmSplash.IpConexio Is Nothing Then
On Error GoTo nor
      Set frmSplash.IpConexio = New ConexioIp
On Error GoTo 0
      frmSplash.IpConexio.Cfg_AppName = "PosDll"
   End If
   
   frmSplash.IpConexio.Cfg_AppPath = AppPath
   
   frmSplash.IpConexio.Cfg_Llicencia = Lic
   If CDbl(Cnf.llicencia) = 1 Then
      LlicenciaValida = True
   Else
      LlicenciaValida = frmSplash.IpConexio.LlicenciaValida(Lic, NumTel, User, Psw, err)
   End If
   Exit Function
nor:
End
End Function

Function NomContingut(NomTaula As String) As String
    
    Select Case NomTaula
       Case "Creades": NomContingut = "CaixesCongelador"
       Case Else: NomContingut = NomTaula
    End Select
    
End Function

Sub ActualitzaTaula(NomTaula As String)
   Dim sql As String, CampsPerSet  As String, i As Integer, Rs As rdoResultset, e
   Dim Camp As String, res As rdoResultset, c, CampsCreate As String
   
   Camp = "TimeStamp"
   If NomTaula = "DeutesAnticips" Then Camp = "Data"
   If NomTaula = "DeutesAnticipsV2" Then Camp = "Data"

   If Not ExisteixTaula(NomTaula & "_Tmp") Then Exit Sub
   TrigguerEsborraTimeStamp NomTaula
      
   CampsCreate = ""
   If NomTaula = "DeutesAnticipsV2" Then
        Set Rs = Db.OpenResultset("Select TOP 1 * From [DeutesAnticips] ")
   Else
        Set Rs = Db.OpenResultset("Select TOP 1 * From [" & NomTaula & "] ")
   End If
   
   CampsPerSet = ""
   For Each e In Rs.rdoColumns
      If Not (UCase(e.Name) = UCase("TimeStamp") Or UCase(e.Name) = UCase("Id")) Then
         If Len(CampsPerSet) > 0 Then CampsPerSet = CampsPerSet & ","
         CampsPerSet = CampsPerSet & "" & NomTaula & ".[" & e.Name & "]=" & NomTaula & "_Tmp.[" & e.Name & "]"
      End If
         If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
         CampsCreate = CampsCreate & "[" & e.Name & "]"
   Next
   
   sql = "Insert Into " & NomTaula & " "
   If NomTaula = "DeutesAnticipsV2" Then sql = "Insert Into DeutesAnticips  "
      
   sql = sql & " (" & CampsCreate & ") Select " & CampsCreate & " From " & NomTaula & "_Tmp  Where Not Id In "
   If NomTaula = "DeutesAnticipsV2" Then
      sql = sql & "(Select Id From DeutesAnticips ) "
   Else
      sql = sql & "(Select isnull(Id,'') Id From " & NomTaula & ") "
   End If
   
   ExecutaComandaSql sql   ' Posem Els Nous
      

   If NomTaula = "ClientsFinals" Then
      ExecutaComandaSql "update clientsfinals set idexterna = '' where idexterna in (Select idexterna from clientsfinals_tmp)"
      sql = "Update " & NomTaula & " Set " & CampsPerSet & " From " & NomTaula & " Join " & NomTaula & "_Tmp On " & NomTaula & "_Tmp.Id = " & NomTaula & ".Id Where " & NomTaula & "_Tmp.[Id] = " & NomTaula & ".[Id] "
      ExecutaComandaSql sql
   Else
      sql = "Update " & NomTaula & " Set " & CampsPerSet & " From " & NomTaula & " Join " & NomTaula & "_Tmp On " & NomTaula & "_Tmp.Id = " & NomTaula & ".Id Where " & NomTaula & "_Tmp.[" & Camp & "] >= " & NomTaula & ".[" & Camp & "] "
      ExecutaComandaSql sql
   End If
   
   If NomTaula = "Missatges" Then
      sql = "Update Missatges Set Missatges.[DataRebut]=Missatges_Tmp.[DataRebut] From Missatges Join Missatges_Tmp On Missatges_Tmp.Id = Missatges.Id and Missatges.[DataRebut] is null "
      ExecutaComandaSql sql
   End If
   TrigguerCreaTimeStamp NomTaula
   EsborraTaula NomTaula & "_Tmp "

End Sub




Sub SetLastData(D As Date)
   Dim Rs As rdoResultset, Q As rdoQuery
   
   If Not ExisteixTaula("Records") Then ExecutaComandaSql "CREATE TABLE [Records] ([TimeStamp] [datetime] NULL ,[Concepte] [nvarchar] (255) NULL ) ON [PRIMARY]"

   ExecutaComandaSql "Delete Records Where Concepte = 'CorreoRebutFins' "
   ExecutaComandaSql "Delete Records Where Concepte = 'UltimaConnexioBona' "
   
   Set Q = Db.CreateQuery("", "Insert Into Records (Concepte,[TimeStamp]) Values ('CorreoRebutFins',?)")
   Q.rdoParameters(0) = D
   Q.Execute
   
   Set Q = Db.CreateQuery("", "Insert Into Records (Concepte,[TimeStamp]) Values ('UltimaConnexioBona',GetDate())")
   Q.Execute

End Sub

Sub TrasbassaMissatgesInterns()
   Dim Rs As rdoResultset, Un As String, Codi As Double, Accio As String, Cb As String, nom As String, Preu As String, P As Integer, Rs2 As rdoResultset, Q As rdoQuery, familia As String, CalCrear As Boolean, Pro As String
    Dim Idioma As String, iD As String
    
On Error GoTo Fi

   Set Rs = Db.OpenResultset("Select Texte From Missatges Where Desti = 'Hit Systems' ")
   While Not Rs.EOF
      Un = Rs("Texte")
      Accio = Car(Un)
      
      P = InStr(Accio, ":")
      If P > 0 Then Accio = Right(Accio, Len(Accio) - P)
      
      If Accio = "Traduir" Then
         Un = Join(Split(Un, "'"), "´")
         Idioma = Car(Un)
         iD = Car(Un)
         P = InStr(Idioma, ":")
         If P > 0 Then Idioma = Right(Idioma, Len(Idioma) - P)
         P = InStr(iD, ":")
         If P > 0 Then iD = Right(iD, Len(iD) - P)
         Set Rs2 = Db.OpenResultset("select * from hit.dbo.diccionari where idstr='Toc_" & iD & "' and App = 'TOC' and Idioma = '" & Idioma & "' and texteoriginal = '" & iD & "'  ")
         If Rs2.EOF Then
            ExecutaComandaSql "Insert Into hit.dbo.diccionari (id,idstr,app,pagina,idioma,texteoriginal,texte) Values (newid(),'Toc_" & iD & "','TOC','','" & Idioma & "','" & iD & "','" & iD & "')"
         End If
         Rs2.Close
      End If
      
      If Accio = "CrearCodiScanner" Then
         Cb = Car(Un)
         nom = Car(Un)
         Preu = Car(Un)
         P = InStr(Cb, ":")
         If P > 0 Then Cb = Right(Cb, Len(Cb) - P)
         P = InStr(nom, ":")
         If P > 0 Then nom = Right(nom, Len(nom) - P)
         P = InStr(Preu, ":")
         If P > 0 Then Preu = Right(Preu, Len(Preu) - P)
         
         CalCrear = True
         Codi = 1
         Set Q = Db.CreateQuery("", "Select Codi From Articles Where Nom = ? And Preu = ? ")
         Q.rdoParameters(0) = nom
         Q.rdoParameters(1) = Preu
         Set Rs2 = Q.OpenResultset
         If Not Rs2.EOF Then
            CalCrear = False
            Codi = Rs2("Codi")
         End If
         Rs2.Close
         Q.Close
         
         If CalCrear Then
            Codi = 1
            Set Rs2 = Db.OpenResultset("Select Max(Codi) From Articles ")
            If Not Rs2.EOF Then If Not IsNull(Rs2(0)) Then Codi = Rs2(0) + 1
            Rs2.Close
            familia = ""
            Set Rs2 = Db.OpenResultset("Select Familia From Articles Where Nom like '%codi nou%' ")
            If Not Rs2.EOF Then If Not IsNull(Rs2(0)) Then familia = Rs2(0)
            Rs2.Close
         
            Set Q = Db.CreateQuery("", "Insert Into Articles (Codi,CodiGenetic,Nom,Preu,PreuMajor,Desconte,EsSumable,Familia,TipoIva,NoDescontesEspecials) Values (?,?,?,?,?,?,?,?,?,?) ")
            Q.rdoParameters(0) = Codi
            Q.rdoParameters(1) = Codi
            Q.rdoParameters(2) = nom
            Q.rdoParameters(3) = Preu
            Q.rdoParameters(4) = Preu
            Q.rdoParameters(5) = 0
            Q.rdoParameters(6) = 1
            Q.rdoParameters(7) = familia
            Q.rdoParameters(8) = 0
            Q.rdoParameters(9) = 0
            Q.Execute
            Missatges_CalEnviar "Articles", ""
         End If
         
         If Not ExisteixTaula("CodisBarres") Then ExecutaComandaSql "CREATE TABLE [CodisBarres] ([Codi] [nvarchar] (255) NULL ,[Producte] [int] NULL ) "
         ExecutaComandaSql "Delete CodisBarres where Codi = '" & Cb & "' Or Producte = " & Codi
         Set Q = Db.CreateQuery("", "Insert Into CodisBarres (Codi,Producte) Values (?,?) ")
         Q.rdoParameters(0) = Cb
         Q.rdoParameters(1) = Codi
         Q.Execute
         
         Missatges_CalEnviar "CodisBarres", ""
      End If
      
      If Accio = "AsignarCodiScanner" Then
         Cb = Car(Un)
         Pro = Car(Un)
         
         P = InStr(Cb, ":")
         If P > 0 Then Cb = Right(Cb, Len(Cb) - P)
         P = InStr(Pro, ":")
         If P > 0 Then Pro = Right(Pro, Len(Pro) - P)
         
         If Not ExisteixTaula("CodisBarres") Then ExecutaComandaSql "CREATE TABLE [CodisBarres] ([Codi] [nvarchar] (255) NULL ,[Producte] [int] NULL ) "
         
         ExecutaComandaSql "Delete CodisBarres where Codi = '" & Cb & "' Or Producte = " & Pro
         Set Q = Db.CreateQuery("", "Insert Into CodisBarres (Codi,Producte) Values (?,?) ")
         Q.rdoParameters(0) = Cb
         Q.rdoParameters(1) = Pro
         Q.Execute
         
         Missatges_CalEnviar "CodisBarres", ""
      End If
      
      Rs.MoveNext
   Wend
   Rs.Close
   
   ExecutaComandaSql "Delete Missatges Where Desti = 'Hit Systems' "
Fi:

End Sub

Sub TrigguerCreaTimeStamp(NomTaula As String)
   Dim sql As String
      
   sql = "CREATE TRIGGER [Ts" & NomTaula & "] ON [" & NomTaula & "] "
   sql = sql & "FOR UPDATE,INSERT AS "
   sql = sql & "Update [" & NomTaula & "] "
   sql = sql & "Set [TimeStamp] = GetDate()"
   sql = sql & "Where Id In (Select Id From Inserted) "
   ExecutaComandaSql sql

End Sub

Sub InformaEstat(Estat As Label, s As String, Optional Mantinguent As Boolean = False)
   Static Last As String
   Static K As Integer
   Static lastD As Date
   Dim st As String
   
   If Not Estat Is Nothing Then
      If Not Mantinguent Then
         Last = s
         st = s
         lastD = Now
      Else
         st = Last
         If DateDiff("s", lastD, Now) > 1 Then
            lastD = Now
            K = (K + 1) Mod 10
            If K = 0 Then st = " - " & st
            If K = 1 Then st = " / " & st
            If K = 2 Then st = " - " & st
            If K = 3 Then st = " \ " & st
            If K = 4 Then st = " | " & st
            If K = 5 Then st = " / " & st
            If K = 6 Then st = " - " & st
            If K = 7 Then st = " \ " & st
            If K = 8 Then st = " | " & st
            If K = 9 Then st = " / " & st
         Else
            lastD = Now
            Exit Sub
         End If
      End If
      Estat.Caption = st
      My_DoEvents
   End If

End Sub

Sub FiltraRegistresImportatsOld(Estat As Label, TaulaOrigen As String)
    Dim dia() As String, Rs As rdoResultset, i As Integer, sql As String, SqlBase As String, CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String
    Dim Ce As String
    ReDim dia(0)
    Dim aaa As String
    Dim DiaDb  As String
    InformaEstat Estat, "Filtran Facturació"
    
    EsborraTaula TaulaOrigen & "2"
    CreaTaulaServit TaulaOrigen & "2", False
    ExecutaComandaSql "ALTER TABLE " & TaulaOrigen & "2 ADD DiaDesti VarChar(255)  NULL"
    
    EsborraTaula TaulaOrigen & "1"
    ExecutaComandaSql "CREATE TABLE [" & TaulaOrigen & "1" & "] ([Id] [uniqueidentifier],[TimeStamp]         [datetime])"
    ExecutaComandaSql "Insert " & TaulaOrigen & "1 Select Id As Id,MAX([TimeStamp]) As [TimeStamp] FROM " & TaulaOrigen & " GROUP BY Id "

    sql = "Insert " & TaulaOrigen & "2  "
    sql = sql & "Select "
    sql = sql & "" & TaulaOrigen & ".Id       As Id ,"
    sql = sql & "" & TaulaOrigen & ".[TimeStamp] As [TimeStamp] ,"
    sql = sql & "QuiStamp As QuiStamp,"
    sql = sql & "Client As Client,"
    sql = sql & "CodiArticle As CodiArticle,"
    sql = sql & "PluUtilitzat As PluUtilitzat,"
    sql = sql & "Viatge As Viatge,"
    sql = sql & "Equip As Equip,"
    sql = sql & "QuantitatDemanada As QuantitatDemanada,"
    sql = sql & "QuantitatTornada As QuantitatTornada,"
    sql = sql & "QuantitatServida as QuantitatServida,"
    sql = sql & "MotiuModificacio As MotiuModificacio,"
    sql = sql & "Hora As Hora,"
    sql = sql & "TipusComanda as TipusComanda,"
    sql = sql & "Comentari As Comentari,"
    sql = sql & "ComentariPer As ComentariPer,"
    sql = sql & "Atribut As Atribut,"
    sql = sql & "CitaDemanada As CitaDemanada,"
    sql = sql & "CitaServida As CitaServida,"
    sql = sql & "CitaTornada As CitaTornada,"
    sql = sql & "DiaDesti As DiaDesti "
    sql = sql & "FROM " & TaulaOrigen & " Join " & TaulaOrigen & "1 On " & TaulaOrigen & ".Id = " & TaulaOrigen & "1.Id And " & TaulaOrigen & ".[TimeStamp] = " & TaulaOrigen & "1.[TimeStamp] "
    ExecutaComandaSql sql
    TaulaOrigen = TaulaOrigen & "2"
    
    Set Rs = Db.OpenResultset("Select Distinct DiaDesti From [" & TaulaOrigen & "] ")
    While Not Rs.EOF
       ReDim Preserve dia(UBound(dia) + 1)
       dia(UBound(dia)) = Rs(0)
       Rs.MoveNext
    Wend
    Rs.Close
    
    CarregaPertinenses CondicioEnviamentClient, CondicioEnviamentViatge, CondicioEnviamentEquip

    For i = 1 To UBound(dia)
       Dim a As String
       InformaEstat Estat, "", True
       InformaEstat Estat, dia(i), True
       
       a = dia(i)
       DiaDb = a
       a = "Servit-" & Mid(a, 3, 2) & "-" & Mid(a, 6, 2) & "-" & Mid(a, 9, 2) & ""
       aaa = a
       If Not ExisteixTaula(a) Then CreaTaulaServit a
       
'If UCase(EmpresaActual) <> UCase("iblatpa") Or UCase(EmpresaActual) = UCase("iartpa") Then
       PauseTrigger NomTaulaData(a), True
       
       ExecutaComandaSql "CREATE NONCLUSTERED INDEX [" & a & "0] ON [dbo].[" & a & "]([id])"
       InformaEstat Estat, "", True
       a = "[" & a & "]"

'***************** Actualitzem Els Ids modificats per copies repetides.
       SqlBase = ""
       SqlBase = SqlBase & "update t set t.Id = s.Id "
       SqlBase = SqlBase & "from " & TaulaOrigen & " t join " & a & " s on s.client = t.client  and s.CodiArticle = t.CodiArticle and s.Viatge = t.Viatge and s.Equip = t.Equip and s.TipusComanda = t.TipusComanda "
       SqlBase = SqlBase & "where DiaDesti = '" & DiaDb & "' "
       ExecutaComandaSql SqlBase
'***************** Actualitzem Els Modificats
' Les comandes dels meus Generadors
       SqlBase = ""
       SqlBase = SqlBase & "Update " & a & " "
       SqlBase = SqlBase & "Set "
       SqlBase = SqlBase & "" & a & ".[TimeStamp] = [" & TaulaOrigen & "].[TimeStamp],"
       SqlBase = SqlBase & "" & a & ".[QuiStamp] = [" & TaulaOrigen & "].[QuiStamp],"
       SqlBase = SqlBase & "" & a & ".[Client] = [" & TaulaOrigen & "].[Client],"
       SqlBase = SqlBase & "" & a & ".[CodiArticle] = [" & TaulaOrigen & "].[CodiArticle],"
       SqlBase = SqlBase & "" & a & ".[PluUtilitzat] = [" & TaulaOrigen & "].[PluUtilitzat],"
'      SqlBase = SqlBase & "" & a & ".[Viatge] = [" & TaulaOrigen & "].[Viatge],"
'      SqlBase = SqlBase & "" & a & ".[Equip] = [" & TaulaOrigen & "].[Equip],"
'       If UCase(EmpresaActual) = UCase("Carne") Then
'           SqlBase = SqlBase & "" & a & ".[QuantitatDemanada] = [" & TaulaOrigen & "].[QuantitatServida],"
'       End If
'      SqlBase = SqlBase & "" & a & ".[QuantitatTornada] = [" & TaulaOrigen & "].[QuantitatTornada],"
       SqlBase = SqlBase & "" & a & ".[QuantitatServida] = [" & TaulaOrigen & "].[QuantitatServida],"
       SqlBase = SqlBase & "" & a & ".[MotiuModificacio] = [" & TaulaOrigen & "].[MotiuModificacio],"
       SqlBase = SqlBase & "" & a & ".[Hora] = [" & TaulaOrigen & "].[Hora],"
       SqlBase = SqlBase & "" & a & ".[TipusComanda] = [" & TaulaOrigen & "].[TipusComanda],"
       SqlBase = SqlBase & "" & a & ".[Comentari] = [" & TaulaOrigen & "].[Comentari],"
       SqlBase = SqlBase & "" & a & ".[ComentariPer] = [" & TaulaOrigen & "].[ComentariPer],"
       SqlBase = SqlBase & "" & a & ".[Atribut] = [" & TaulaOrigen & "].[Atribut],"
       SqlBase = SqlBase & "" & a & ".[CitaDemanada] = [" & TaulaOrigen & "].[CitaDemanada]"
'       SqlBase = SqlBase & "" & a & ".[CitaTornada] = [" & TaulaOrigen & "].[CitaTornada],"
'       SqlBase = SqlBase & "" & a & ".[CitaServida] = [" & TaulaOrigen & "].[CitaServida],"
       SqlBase = SqlBase & "From " & TaulaOrigen & " "
       SqlBase = SqlBase & "Join " & a & " "
       SqlBase = SqlBase & "On isnull(" & a & ".Hora,0) = 92 And  " & a & ".Id = [" & TaulaOrigen & "].Id And [" & TaulaOrigen & "].DiaDesti = '" & DiaDb & "' "
'       SqlBase = SqlBase & "Where " & a & ".[TimeStamp] <= [" & TaulaOrigen & "].[TimeStamp]"
'********* De moment descartem el timestamp doncs a la botiga no es modifica !!!!

' Per Concepte de Viatge
'       If Len(CondicioEnviamentViatge) > 0 Then
'          Sql = SqlBase & " And " & a & ".Viatge In(" & CondicioEnviamentViatge & ") "
          sql = SqlBase
          ExecutaComandaSql SqlBase
'       End If
       InformaEstat Estat, "", True
       

' Els nous ke no hagin arribat   92
       
       SqlBase = ""
       SqlBase = SqlBase & "Update " & a & " "
       SqlBase = SqlBase & "Set "
       SqlBase = SqlBase & "" & a & ".[TimeStamp] = [" & TaulaOrigen & "].[TimeStamp],"
       SqlBase = SqlBase & "" & a & ".[QuiStamp] = [" & TaulaOrigen & "].[QuiStamp],"
       SqlBase = SqlBase & "" & a & ".[Client] = [" & TaulaOrigen & "].[Client],"
       SqlBase = SqlBase & "" & a & ".[CodiArticle] = [" & TaulaOrigen & "].[CodiArticle],"
       SqlBase = SqlBase & "" & a & ".[PluUtilitzat] = [" & TaulaOrigen & "].[PluUtilitzat],"
       SqlBase = SqlBase & "" & a & ".[Viatge] = [" & TaulaOrigen & "].[Viatge],"
       SqlBase = SqlBase & "" & a & ".[Equip] = [" & TaulaOrigen & "].[Equip],"
'      If UCase(EmpresaActual) = UCase("Carne") Then
'         SqlBase = SqlBase & "" & a & ".[QuantitatDemanada] = [" & TaulaOrigen & "].[QuantitatServida],"
'      End If
       SqlBase = SqlBase & "" & a & ".[QuantitatTornada] = [" & TaulaOrigen & "].[QuantitatTornada],"
       SqlBase = SqlBase & "" & a & ".[QuantitatServida] = [" & TaulaOrigen & "].[QuantitatServida],"
       SqlBase = SqlBase & "" & a & ".[MotiuModificacio] = [" & TaulaOrigen & "].[MotiuModificacio],"
       SqlBase = SqlBase & "" & a & ".[Hora] = [" & TaulaOrigen & "].[Hora],"
       SqlBase = SqlBase & "" & a & ".[TipusComanda] = [" & TaulaOrigen & "].[TipusComanda],"
       SqlBase = SqlBase & "" & a & ".[Comentari] = [" & TaulaOrigen & "].[Comentari],"
       SqlBase = SqlBase & "" & a & ".[ComentariPer] = [" & TaulaOrigen & "].[ComentariPer],"
       SqlBase = SqlBase & "" & a & ".[Atribut] = [" & TaulaOrigen & "].[Atribut]"
'       SqlBase = SqlBase & "" & a & ".[CitaDemanada] = [" & TaulaOrigen & "].[CitaDemanada]"
'       SqlBase = SqlBase & "" & a & ".[CitaTornada] = [" & TaulaOrigen & "].[CitaTornada],"
'       SqlBase = SqlBase & "" & a & ".[CitaServida] = [" & TaulaOrigen & "].[CitaServida],"
       SqlBase = SqlBase & "From " & TaulaOrigen & " "
       SqlBase = SqlBase & "Join " & a & " "
       SqlBase = SqlBase & "On isnull(" & a & ".Hora,0) <> 92 And  " & a & ".Id =  [" & TaulaOrigen & "].Id And [" & TaulaOrigen & "].DiaDesti = '" & DiaDb & "' "
'       SqlBase = SqlBase & "Where " & a & ".[TimeStamp] <= [" & TaulaOrigen & "].[TimeStamp]"
'********* De moment descartem el timestamp doncs a la botiga no es modifica !!!!

' Per Concepte de Viatge
'       If Len(CondicioEnviamentViatge) > 0 Then
'          Sql = SqlBase & " And " & a & ".Viatge In(" & CondicioEnviamentViatge & ") "
          sql = SqlBase
          ExecutaComandaSql SqlBase
'       End If
       InformaEstat Estat, "", True
       
       
'' Els Meus CLients (Facturació)
'       SqlBase = ""
'       SqlBase = SqlBase & "Update " & a & " "
'       SqlBase = SqlBase & "Set "
'       SqlBase = SqlBase & "" & a & ".[TimeStamp] = [" & TaulaOrigen & "].[TimeStamp],"
'       SqlBase = SqlBase & "" & a & ".[QuiStamp] = [" & TaulaOrigen & "].[QuiStamp],"
'       SqlBase = SqlBase & "" & a & ".[Client] = [" & TaulaOrigen & "].[Client],"
'       SqlBase = SqlBase & "" & a & ".[CodiArticle] = [" & TaulaOrigen & "].[CodiArticle],"
'       SqlBase = SqlBase & "" & a & ".[PluUtilitzat] = [" & TaulaOrigen & "].[PluUtilitzat],"
'       SqlBase = SqlBase & "" & a & ".[Viatge] = [" & TaulaOrigen & "].[Viatge],"
'       SqlBase = SqlBase & "" & a & ".[Equip] = [" & TaulaOrigen & "].[Equip],"
'       SqlBase = SqlBase & "" & a & ".[QuantitatDemanada] = [" & TaulaOrigen & "].[QuantitatDemanada],"
'       SqlBase = SqlBase & "" & a & ".[QuantitatTornada] = [" & TaulaOrigen & "].[QuantitatTornada],"
'       SqlBase = SqlBase & "" & a & ".[QuantitatServida] = [" & TaulaOrigen & "].[QuantitatServida],"
'       SqlBase = SqlBase & "" & a & ".[MotiuModificacio] = [" & TaulaOrigen & "].[MotiuModificacio],"
'       SqlBase = SqlBase & "" & a & ".[Hora] = [" & TaulaOrigen & "].[Hora],"
'       SqlBase = SqlBase & "" & a & ".[TipusComanda] = [" & TaulaOrigen & "].[TipusComanda],"
'       SqlBase = SqlBase & "" & a & ".[Comentari] = [" & TaulaOrigen & "].[Comentari],"
'       SqlBase = SqlBase & "" & a & ".[ComentariPer] = [" & TaulaOrigen & "].[ComentariPer],"
'       SqlBase = SqlBase & "" & a & ".[Atribut] = [" & TaulaOrigen & "].[Atribut],"
'       SqlBase = SqlBase & "" & a & ".[CitaDemanada] = [" & TaulaOrigen & "].[CitaDemanada],"
'       SqlBase = SqlBase & "" & a & ".[CitaTornada] = [" & TaulaOrigen & "].[CitaTornada],"
'       SqlBase = SqlBase & "" & a & ".[CitaServida] = [" & TaulaOrigen & "].[CitaServida]"
'       SqlBase = SqlBase & "From " & TaulaOrigen & " "
'       SqlBase = SqlBase & "Join " & a & " "
'       SqlBase = SqlBase & "On " & a & ".Id = [" & TaulaOrigen & "].Id "
'       SqlBase = SqlBase & "And [" & TaulaOrigen & "].DiaDesti = '" & DiaDb & "' "
'       SqlBase = SqlBase & "Where " & a & ".[TimeStamp] <= [" & TaulaOrigen & "].[TimeStamp]"
'
''       If Len(CondicioEnviamentClient) > 0 Then
''          Sql = SqlBase & " And " & a & ".Client In(" & CondicioEnviamentClient & ") "
''          If Len(CondicioEnviamentViatge) > 0 Then ' I que no es reparteixi desde aqui
''             Sql = SqlBase & " And Not " & a & ".Viatge In(" & CondicioEnviamentViatge & ") "
''          End If
'          ExecutaComandaSql SqlBase
''       End If
'       InformaEstat Estat, "", True
       
' Posem Els Nous
       sql = ""
       sql = sql & "Insert Into " & a & " ([Id],[TimeStamp],[QuiStamp],[Client],[CodiArticle],[PluUtilitzat],[Viatge],[Equip],[QuantitatDemanada],[QuantitatTornada],[QuantitatServida],[MotiuModificacio],[Hora],[TipusComanda],[Comentari],[ComentariPer],[Atribut],[CitaDemanada],[CitaServida],[CitaTornada]) "
       sql = sql & "Select "
       sql = sql & "[Id],"
       sql = sql & "[TimeStamp],"
       sql = sql & "[QuiStamp],"
       sql = sql & "[Client],"
       sql = sql & "[CodiArticle],"
       sql = sql & "[PluUtilitzat],"
       sql = sql & "[Viatge],"
       sql = sql & "[Equip],"
       sql = sql & "[QuantitatDemanada],"
       sql = sql & "[QuantitatTornada],"
       sql = sql & "[QuantitatServida],"
       sql = sql & "[MotiuModificacio],"
       sql = sql & "71,"
       sql = sql & "[TipusComanda],"
       sql = sql & "[Comentari],"
       sql = sql & "[ComentariPer],"
       sql = sql & "[Atribut],"
       sql = sql & "[CitaDemanada],"
       sql = sql & "[CitaServida],"
       sql = sql & "[CitaTornada]"
       sql = sql & "From " & TaulaOrigen & " "
       sql = sql & "Where "
       sql = sql & "[" & TaulaOrigen & "].DiaDesti = '" & DiaDb & "' "
       sql = sql & "And Not [" & TaulaOrigen & "].Id In (Select Id From " & a & ")"
       ExecutaComandaSql sql
       
'       Ce = ""
'       If Len(CondicioEnviamentClient) > 0 Then Ce = Ce & "Client In(" & CondicioEnviamentClient & ") "
'       If Len(CondicioEnviamentViatge) > 0 And Len(Ce) > 0 Then Ce = Ce & " Or "
'       If Len(CondicioEnviamentViatge) > 0 Then Ce = Ce & "Viatge In(" & CondicioEnviamentViatge & ") "
'       If Len(CondicioEnviamentEquip) > 0 And Len(Ce) > 0 Then Ce = Ce & " Or "
'       If Len(CondicioEnviamentEquip) > 0 Then Ce = Ce & "Equip In(" & CondicioEnviamentEquip & " ) "
'       If Len(Ce) > 0 Then
'          Sql = Sql & " And (" & Ce & ")"
'          ExecutaComandaSql Sql
'          Debug.Print Sql
'       End If
'       InformaEstat Estat, "", True
       
       PauseTrigger NomTaulaData(aaa), False
       InformaEstat Estat, "", True
    Next
        
End Sub
Sub FiltraRegistresImportats(Estat As Label, TaulaOrigen As String)
    Dim Rs As rdoResultset, rsRebut As rdoResultset, rsGraella As rdoResultset, rsVE As rdoResultset
    Dim i As Integer, sql As String, dia() As String
    Dim iD As String, Viatge As String, equip As String, article As String, client As String
    Dim TaulaServit As String
    Dim Demanat As Boolean
    
    Demanat = False
    If UCase(EmpresaActual) = UCase("Pa Natural") Then Demanat = True
    
    InformaEstat Estat, "Filtran Facturació"
    
    ReDim dia(0)
    Set Rs = Db.OpenResultset("Select Distinct DiaDesti From [" & TaulaOrigen & "] ")
    While Not Rs.EOF
       ReDim Preserve dia(UBound(dia) + 1)
       dia(UBound(dia)) = Rs(0)
       Rs.MoveNext
    Wend
    Rs.Close

    For i = 1 To UBound(dia)
       Dim a As String
       InformaEstat Estat, "", True
       InformaEstat Estat, dia(i), True
       
       TaulaServit = "Servit-" & Mid(dia(i), 3, 2) & "-" & Mid(dia(i), 6, 2) & "-" & Mid(dia(i), 9, 2)
       If Not ExisteixTaula(TaulaServit) Then CreaTaulaServit TaulaServit
       
       PauseTrigger NomTaulaData(TaulaServit), True
       
       Set rsRebut = Db.OpenResultset("select * from " & TaulaOrigen & " where diaDesti='" & dia(i) & "'")
       While Not rsRebut.EOF
            iD = rsRebut("id")
            client = 0
            If Not IsNull(rsRebut("client")) Then client = rsRebut("client")
            article = rsRebut("CodiArticle")
            Viatge = rsRebut("viatge")
            equip = rsRebut("equip")

            If UCase(EmpresaActual) = UCase("SISTARE") Then
                Set rsVE = Db.OpenResultset("select Top 1 isnull(Viatge,'') Viatge, isnull(Equip,'') Equip from comandesmemotecnicperclient where codiarticle = '" & article & "' and (client = '" & client & "' or client is null) order by timestamp desc ")
                If Not rsVE.EOF Then
                    Viatge = rsVE("Viatge")
                    equip = rsVE("Equip")
                End If
            End If
            
            If Left(iD, 5) = Right("00000" & client, 5) Then 'Es un articulo que no estaba en la graella original
                Set rsGraella = Db.OpenResultset("select * from [" & TaulaServit & "] where cast(id as varchar(2000))='" & iD & "' or comentari like '%IdOld:" & iD & "%'")
            Else
                Set rsGraella = Db.OpenResultset("select * from [" & TaulaServit & "] where id='" & iD & "'")
            End If
            
            If Not rsGraella.EOF Then 'Ya existe el id en Servit
                sql = "update [" & TaulaServit & "] set "
                sql = sql & "[" & TaulaServit & "].[QuantitatServida] = " & rsRebut("QuantitatServida") & " "
                If Demanat Then
                    sql = sql & ", [" & TaulaServit & "].[QuantitatDemanada] = " & rsRebut("QuantitatServida") & " "
                End If
                sql = sql & "where id='" & rsGraella("Id") & "'"
            Else
                Set rsGraella = Db.OpenResultset("select * from [" & TaulaServit & "] where viatge='" & Viatge & "' and equip='" & equip & "' and codiArticle='" & article & "' and client='" & client & "'")
                If Not rsGraella.EOF Then 'Ya existe un registro con mismo viaje, equipo, articulo i cliente
                    sql = "update [" & TaulaServit & "] set "
                    sql = sql & "[" & TaulaServit & "].[QuantitatServida] = " & rsRebut("QuantitatServida") & " "
                    If Demanat Then
                        sql = sql & ", [" & TaulaServit & "].[QuantitatDemanada] = " & rsRebut("QuantitatServida") & " "
                    End If
                    sql = sql & "where viatge='" & Viatge & "' and equip='" & equip & "' and codiArticle='" & article & "'"
                    sql = sql & " and client = '" & client & "'"
                Else 'No hay registro en Servit
                    Set rsVE = Db.OpenResultset("select Top 1 isnull(Viatge,'') Viatge, isnull(Equip,'') Equip from comandesmemotecnicperclient where codiarticle = '" & article & "' and (client = '" & client & "' or client is null) order by timestamp desc ")
                    If Not rsVE.EOF Then
                        Viatge = rsVE("Viatge")
                        equip = rsVE("Equip")
                    End If
                
                    If Left(iD, 5) = Right("00000" & client, 5) Then
                        sql = "insert into [" & TaulaServit & "] (Id, TimeStamp, QuiStamp, Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada)"
                        If Demanat Then
                            sql = sql & "select newid(), DATEADD(m,-1 ,TimeStamp), QuiStamp, Client, CodiArticle, PluUtilitzat, '" & Viatge & "', '" & equip & "', QuantitatServida, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari+'[IdOld:" & iD & "]', ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada "
                        Else
                            sql = sql & "select newid(), DATEADD(m,-1 ,TimeStamp), QuiStamp, Client, CodiArticle, PluUtilitzat, '" & Viatge & "', '" & equip & "', QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari+'[IdOld:" & iD & "]', ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada "
                        End If
                        sql = sql & "from [" & TaulaOrigen & "] "
                        sql = sql & "where id='" & iD & "'"
                    Else
                        sql = "insert into [" & TaulaServit & "] (Id, TimeStamp, QuiStamp, Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada)"
                        If Demanat Then
                            sql = sql & "select Id, DATEADD(m,-1 ,TimeStamp), QuiStamp, Client, CodiArticle, PluUtilitzat, '" & Viatge & "', '" & equip & "', QuantitatServida, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada "
                        Else
                            sql = sql & "select Id, DATEADD(m,-1 ,TimeStamp), QuiStamp, Client, CodiArticle, PluUtilitzat, '" & Viatge & "', '" & equip & "', QuantitatDemanada, QuantitatTornada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut, CitaDemanada, CitaServida, CitaTornada "
                        End If
                        sql = sql & "from [" & TaulaOrigen & "] "
                        sql = sql & "where id='" & iD & "'"
                    End If
                End If
            End If
            ExecutaComandaSql sql
            rsRebut.MoveNext
       Wend
       
       PauseTrigger NomTaulaData(TaulaServit), False
       
       InformaEstat Estat, "", True
    Next
        
End Sub

Sub CreaTaulaServit(NomTaula As String, Optional CreaTriguer As Boolean = True)
    Dim i As Integer, Trobada As Boolean, sql As String, CalArreglar As Boolean
    
    CalArreglar = False
    If ExisteixTaula(NomTaula) And Not ExisteixTaula(NomTaula & "Trace") Then CalArreglar = True
   
    sql = "CREATE TABLE [" & NomTaula & "] ("
    If NomTaula = "Servit_Temporal" Then
        sql = sql & "[Id] [nvarchar] (255)  DEFAULT (newid()), "
    Else
        sql = sql & "[Id]                [uniqueidentifier] Default (NEWID()),"
    End If
    
   sql = sql & "[TimeStamp]         [datetime]              Default (GetDate())  ,"
   sql = sql & "[QuiStamp]          [nvarchar] (255)        Default (Host_Name()),"
   sql = sql & "[Client]            [float]          Null ,"
   sql = sql & "[CodiArticle]       [int]            Null ,"
   sql = sql & "[PluUtilitzat]      [nvarchar] (255) Null ,"
   sql = sql & "[Viatge]            [nvarchar] (255) Null ,"
   sql = sql & "[Equip]             [nvarchar] (255) Null ,"
   sql = sql & "[QuantitatDemanada] [float]                 Default (0),"
   sql = sql & "[QuantitatTornada]  [float]                 Default (0),"
   sql = sql & "[QuantitatServida]  [float]                 Default (0),"
   sql = sql & "[MotiuModificacio]  [nvarchar] (255) Null ,"
   sql = sql & "[Hora]              [float]          Null ,"
   sql = sql & "[TipusComanda]      [float]          Null ,"
   sql = sql & "[Comentari]         [nvarchar] (255) Null ,"
   sql = sql & "[ComentariPer]      [nvarchar] (255) Null ,"
   sql = sql & "[Atribut]           [Int]            Null ,"
   sql = sql & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaServida]       [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaTornada]       [nvarchar] (255)        Default ('') "
   sql = sql & ") ON [PRIMARY]"
   ExecutaComandaSql (sql)
   
   sql = "CREATE TABLE [" & NomTaula & "Reg] ("
   sql = sql & "[Modificat]         [datetime]              Default (GetDate())  ,"
   sql = sql & "[Id]                [nvarchar] (255) Null ,"
   sql = sql & "[TimeStamp]         [datetime]       Null ,"
   sql = sql & "[QuiStamp]          [nvarchar] (255) Null ,"
   sql = sql & "[Client]            [float]          Null ,"
   sql = sql & "[CodiArticle]       [int]            Null ,"
   sql = sql & "[PluUtilitzat]      [nvarchar] (255) Null ,"
   sql = sql & "[Viatge]            [nvarchar] (255) Null ,"
   sql = sql & "[Equip]             [nvarchar] (255) Null ,"
   sql = sql & "[QuantitatDemanada] [float]                 Default (0),"
   sql = sql & "[QuantitatTornada]  [float]                 Default (0),"
   sql = sql & "[QuantitatServida]  [float]                 Default (0),"
   sql = sql & "[MotiuModificacio]  [nvarchar] (255) Null ,"
   sql = sql & "[Hora]              [float]          Null ,"
   sql = sql & "[TipusComanda]      [float]          Null ,"
   sql = sql & "[Comentari]         [nvarchar] (255) Null ,"
   sql = sql & "[ComentariPer]      [nvarchar] (255) Null ,"
   sql = sql & "[Atribut]           [Int]            Null ,"
   sql = sql & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaServida]       [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaTornada]       [nvarchar] (255)        Default ('') "
   sql = sql & ") ON [PRIMARY]"
   ExecutaComandaSql (sql)
   
   If CreaTriguer Then
   sql = "CREATE TRIGGER [M_" & NomTaula & "] ON [" & NomTaula & "] "
   sql = sql & "AFTER INSERT,UPDATE,DELETE AS "
   sql = sql & "Update [" & NomTaula & "] Set [TimeStamp] = GetDate(),    [QuiStamp]  = Host_Name() Where Id In (Select Id From Inserted) "
   sql = sql & "Insert Into ComandesModificades Select Id As Id,GetDate() As [TimeStamp],'" & NomTaula & "' As TaulaOrigen From Inserted "
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from [" & NomTaula & "] Where Id In (Select Id From Inserted)"
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp]+'BORRAT!!!',Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from deleted Where not Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatTornada)  Update [" & NomTaula & "] Set [CitaTornada]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatTornada  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaTornada]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatServida)  Update [" & NomTaula & "] Set [CitaServida]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatServida  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaServida]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatDemanada) Update [" & NomTaula & "] Set [CitaDemanada] = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatDemanada AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaDemanada] AS VarChar(255)) Where Id In (Select Id From Inserted) "
      ExecutaComandaSql (sql)
   End If

    If CalArreglar Then CreaTaulaServitArregla NomTaula
    
End Sub

Function NomTaulaLikes(dia As Date) As String
    Dim i As Integer, Trobada As Boolean, sql As String, CalArreglar As Boolean

    NomTaulaLikes = "Likes_" & Year(dia) & "-" & Month(dia)
    ExecutaComandaSql ("Drop table [" & NomTaulaLikes & "]")
    
    If Not ExisteixTaula(NomTaulaLikes) Then
        sql = "CREATE TABLE [" & NomTaulaLikes & "] ("
        sql = sql & "[TimeStamp]         [datetime]              Default (GetDate())  ,"
        sql = sql & "[QuiStamp]          [nvarchar] (255)        Default (Host_Name()),"
        sql = sql & "[Que]               [nvarchar] (255) Null ,"
        sql = sql & "[Origen]            [nvarchar] (255) Null ,"
        sql = sql & "[Destino]           [nvarchar] (255) Null ,"
        sql = sql & "[Nota]              [nvarchar] (255) Null "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql (sql)
    End If
   
    
End Function


Sub CreaTaulaEstimat(NomTaula As String)
    Dim i As Integer, Trobada As Boolean, sql As String, CalArreglar As Boolean
    
   sql = "CREATE TABLE [" & NomTaula & "] ("
   sql = sql & "[Id] [nvarchar] (255)  DEFAULT (newid()), "
   sql = sql & "[TimeStamp]         [datetime]              Default (GetDate())  ,"
   sql = sql & "[QuiStamp]          [nvarchar] (255)        Default (Host_Name()),"
   sql = sql & "[Client]            [float]          Null ,"
   sql = sql & "[CodiArticle]       [int]            Null ,"
   sql = sql & "[PluUtilitzat]      [nvarchar] (255) Null ,"
   sql = sql & "[Viatge]            [nvarchar] (255) Null ,"
   sql = sql & "[Equip]             [nvarchar] (255) Null ,"
   sql = sql & "[QuantitatDemanada] [float]                 Default (0),"
   sql = sql & "[QuantitatTornada]  [float]                 Default (0),"
   sql = sql & "[QuantitatServida]  [float]                 Default (0),"
   sql = sql & "[MotiuModificacio]  [nvarchar] (255) Null ,"
   sql = sql & "[Hora]              [float]          Null ,"
   sql = sql & "[TipusComanda]      [float]          Null ,"
   sql = sql & "[Comentari]         [nvarchar] (255) Null ,"
   sql = sql & "[ComentariPer]      [nvarchar] (255) Null ,"
   sql = sql & "[Atribut]           [Int]            Null ,"
   sql = sql & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaServida]       [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaTornada]       [nvarchar] (255)        Default ('') "
   sql = sql & ") ON [PRIMARY]"
   ExecutaComandaSql (sql)
   
End Sub


Sub CreaTaulaServit2(NomTaula As String)
Dim sql, CreaTriguer, CalArreglar

   sql = "CREATE TABLE [" & NomTaula & "] ("
   sql = sql & "[Id]                [nvarchar] (255) Default (NEWID()),"
   sql = sql & "[TimeStamp]         [datetime]              Default (GetDate())  ,"
   sql = sql & "[QuiStamp]          [nvarchar] (255)        Default (Host_Name()),"
   sql = sql & "[Client]            [float]          Null ,"
   sql = sql & "[CodiArticle]       [int]            Null ,"
   sql = sql & "[PluUtilitzat]      [nvarchar] (255) Null ,"
   sql = sql & "[Viatge]            [nvarchar] (255) Null ,"
   sql = sql & "[Equip]             [nvarchar] (255) Null ,"
   sql = sql & "[QuantitatDemanada] [float]                 Default (0),"
   sql = sql & "[QuantitatTornada]  [float]                 Default (0),"
   sql = sql & "[QuantitatServida]  [float]                 Default (0),"
   sql = sql & "[MotiuModificacio]  [nvarchar] (255) Null ,"
   sql = sql & "[Hora]              [float]          Null ,"
   sql = sql & "[TipusComanda]      [float]          Null ,"
   sql = sql & "[Comentari]         [nvarchar] (255) Null ,"
   sql = sql & "[ComentariPer]      [nvarchar] (255) Null ,"
   sql = sql & "[Atribut]           [Int]            Null ,"
   sql = sql & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaServida]       [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaTornada]       [nvarchar] (255)        Default ('') "
   sql = sql & ") ON [PRIMARY]"
   ExecutaComandaSql (sql)
   
   If CreaTriguer Then
   sql = "CREATE TRIGGER [M_" & NomTaula & "] ON [" & NomTaula & "] "
   sql = sql & "AFTER INSERT,UPDATE,DELETE AS "
   sql = sql & "Update [" & NomTaula & "] Set [TimeStamp] = GetDate(),    [QuiStamp]  = Host_Name() Where Id In (Select Id From Inserted) "
   sql = sql & "Insert Into ComandesModificades Select Id As Id,GetDate() As [TimeStamp],'" & NomTaula & "' As TaulaOrigen From Inserted "
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from [" & NomTaula & "] Where Id In (Select Id From Inserted)"
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp]+'BORRAT!!!',Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from deleted Where not Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatTornada)  Update [" & NomTaula & "] Set [CitaTornada]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatTornada  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaTornada]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatServida)  Update [" & NomTaula & "] Set [CitaServida]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatServida  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaServida]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatDemanada) Update [" & NomTaula & "] Set [CitaDemanada] = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatDemanada AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaDemanada] AS VarChar(255)) Where Id In (Select Id From Inserted) "
      ExecutaComandaSql (sql)
   End If

   CalArreglar = False
   If Not ExisteixTaula(NomTaula & "Trace") Then CalArreglar = True

    If CalArreglar Then CreaTaulaServitArregla NomTaula
    
End Sub






Sub CreaTaulaServitArregla(NomTaula As String)
   Dim i As Integer, Trobada As Boolean, sql As String

   sql = "CREATE TABLE [" & NomTaula & "Trace] ("
   sql = sql & "[Modificat]         [datetime]              Default (GetDate())  ,"
   sql = sql & "[Id]                [nvarchar] (255) Null ,"
   sql = sql & "[TimeStamp]         [datetime]       Null ,"
   sql = sql & "[QuiStamp]          [nvarchar] (255) Null ,"
   sql = sql & "[Client]            [float]          Null ,"
   sql = sql & "[CodiArticle]       [int]            Null ,"
   sql = sql & "[PluUtilitzat]      [nvarchar] (255) Null ,"
   sql = sql & "[Viatge]            [nvarchar] (255) Null ,"
   sql = sql & "[Equip]             [nvarchar] (255) Null ,"
   sql = sql & "[QuantitatDemanada] [float]                 Default (0),"
   sql = sql & "[QuantitatTornada]  [float]                 Default (0),"
   sql = sql & "[QuantitatServida]  [float]                 Default (0),"
   sql = sql & "[MotiuModificacio]  [nvarchar] (255) Null ,"
   sql = sql & "[Hora]              [float]          Null ,"
   sql = sql & "[TipusComanda]      [float]          Null ,"
   sql = sql & "[Comentari]         [nvarchar] (255) Null ,"
   sql = sql & "[ComentariPer]      [nvarchar] (255) Null ,"
   sql = sql & "[Atribut]           [Int]            Null ,"
   sql = sql & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaServida]       [nvarchar] (255)        Default (''),"
   sql = sql & "[CitaTornada]       [nvarchar] (255)        Default ('') "
   sql = sql & ") ON [PRIMARY]"
   ExecutaComandaSql (sql)
   
   sql = "DROP TRIGGER [M_" & NomTaula & "] "
   ExecutaComandaSql (sql)
   
   sql = "CREATE TRIGGER [M_" & NomTaula & "] ON [" & NomTaula & "] "
   sql = sql & "AFTER INSERT,UPDATE,DELETE AS "
   sql = sql & "Update [" & NomTaula & "] Set [TimeStamp] = GetDate(),    [QuiStamp]  = Host_Name() Where Id In (Select Id From Inserted) "
   sql = sql & "Insert Into ComandesModificades Select Id As Id,GetDate() As [TimeStamp],'" & NomTaula & "' As TaulaOrigen From Inserted "
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from [" & NomTaula & "] Where Id In (Select Id From Inserted)"
   sql = sql & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp]+'BORRAT!!!',Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from deleted Where not Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatTornada)  Update [" & NomTaula & "] Set [CitaTornada]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatTornada  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaTornada]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatServida)  Update [" & NomTaula & "] Set [CitaServida]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatServida  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaServida]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
   sql = sql & "If Update (QuantitatDemanada) Update [" & NomTaula & "] Set [CitaDemanada] = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatDemanada AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaDemanada] AS VarChar(255)) Where Id In (Select Id From Inserted) "
   
   ExecutaComandaSql (sql)

End Sub







Sub Interpreta_SqlTrans(Optional Estat As Label = Nothing)
   Dim Fil As String, Camps() As String, Tipus() As String, Valor()  As String
   Dim Ll As String, L As String, f, sql As String, LlistaC As String, LlistaCc As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean
   
   AreglaNomsMalPosats
   
   FtpCopy

   ReDim Camps(0)
   My_DoEvents
   InterpretaFitchersVentes Estat
   Interpreta_SqlTrans_2 "Missatges", Estat
   TrasbassaMissatgesInterns
   Interpreta_SqlTrans_2 "FeinaFeta", Estat
   Interpreta_SqlTrans_2 "Creades", Estat
   Interpreta_SqlTrans_2 "Dependentes", Estat
   Interpreta_SqlTrans_2 "ClientsFinals", Estat
   Interpreta_SqlTrans_2 "ClientsFinalsPropietats", Estat
   Interpreta_SqlTrans_2 "DeutesAnticips", Estat
   Interpreta_SqlTrans_2 "DeutesAnticipsV2", Estat
   Interpreta_SqlTrans_2 "TarifaEspecial*", Estat
   Interpreta_SqlTrans_Tpv Estat
   
   Interpreta_SqlTrans_RecepcioMP Estat
   Interpreta_SqlTrans_Teclat Estat
   Interpreta_SqlTrans_Succeit Estat
   Interpreta_SqlTrans_ComandaRevisada Estat
   Interpreta_SqlTrans_Segurata Estat
   Interpreta_SqlTrans_Comanda Estat
   Interpreta_SqlTrans_Referencies Estat
   Interpreta_SqlTrans_Dedos Estat
   
   If EmpresaActual = "GardenPonc" Then
        Obtencio_Garden_articles_sql
   End If
   
   Fil = Dir(AppPath & "\*appccComo.SqlTrans")
   While Len(Fil) > 0
       MyKill AppPath & "\" & Fil
       Fil = Dir
   Wend
   Fil = Dir(AppPath & "\*appccCuando.SqlTrans")
   While Len(Fil) > 0
       MyKill AppPath & "\" & Fil
       Fil = Dir
   Wend
   Fil = Dir(AppPath & "\*appccTareas.SqlTrans")
   While Len(Fil) > 0
       MyKill AppPath & "\" & Fil
       Fil = Dir
   Wend
   Fil = Dir(AppPath & "\*appccTareasAsignadas.SqlTrans")
   While Len(Fil) > 0
       MyKill AppPath & "\" & Fil
       Fil = Dir
   Wend
   Fil = Dir(AppPath & "\*appccTareasResueltas.SqlTrans")
   While Len(Fil) > 0
       MyKill AppPath & "\" & Fil
       Fil = Dir
   Wend
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*.SqlTrans")
   While Len(Fil) > 0
    If InStr(Fil, "PreferenciasTeclat") = 0 Then
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  ExecutaComandaSql Ll
               End If
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  CalCrearQuery = True
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  sql = "Insert Into " & NomTaula & " (" & LlistaC & ") Values (" & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
               End If
On Error Resume Next
               For i = 0 To UBound(Camps) - 1
                  Q.rdoParameters(i) = BuscaTipus(Q.rdoParameters(i).Type, NormalitzaDes(Car(L)))
               Next
                  Q.Execute
On Error GoTo 0
            End If
         End If
      Wend
err:
      Close #f
      FitcherProcesat Fil
    End If
    Fil = Dir
   Wend

End Sub
Function LlicenciaClient(llicencia As Double) As Integer
   Dim Rs As rdoResultset
   
   LlicenciaClient = 0
   
   
      
   Set Rs = Db.OpenResultset("Select Valor1 From ParamsHw Where Codi = '" & llicencia & "'  And Tipus = 1 ")
   If Not Rs.EOF Then
        If Not IsNull(Rs("Valor1")) Then
            If Rs("Valor1") > 0 And Rs("Valor1") <= 9999 Then
                LlicenciaClient = Rs("Valor1")
            End If
        End If
   End If
   Rs.Close

End Function


Sub Interpreta_SqlTrans_Tpv(Optional Estat As Label = Nothing)
   Dim Fil As String, CalCrearQuery As Boolean, f, L As String, Ll As String, Lll As String, Files() As String, P As Integer, Pp As Integer, llicencia As Double, Variable As String, Valor As String
   Dim Camps() As String, Tipus()  As String, NomTaula As String
   Dim LlistaCc As String, LlistaC As String, i As Integer, sql As String, Q As rdoQuery, CampsPerSet As String, LlistaCreate As String
   Dim FilesAImportar() As String, Kk As Integer
   Dim TeId As Boolean, SomDependentes As Boolean, SomDependentes_K As Integer, SomDependentes_P As Integer
   Dim ik As Integer
   
   Fil = Dir(AppPath & "\*ConfiguracioTpv.Cfg")
   ReDim Camps(0)
   ReDim Tipus(0)
   ReDim FilesAImportar(0)
   While Len(Fil) > 0
      ReDim Preserve FilesAImportar(UBound(FilesAImportar) + 1)
      FilesAImportar(UBound(FilesAImportar)) = Fil
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilesAImportar)
      f = FreeFile
      Fil = FilesAImportar(Kk)
      InformaEstat Estat, "Interpretant : " & Fil
      My_DoEvents
      P = InStr(Fil, "Tpv_Configuracio_")
      llicencia = 0
      If P > 0 Then
         Pp = InStr(P, Fil, "]")
         If Pp > 0 Then
            If IsNumeric(Mid(Fil, P + 17, Pp - P - 17)) Then
               llicencia = Mid(Fil, P + 17, Pp - P - 17)
            Else
               llicencia = 0
            End If
            llicencia = LlicenciaClient(llicencia)
         End If
      End If
      
      If llicencia > 0 Then
         Open AppPath & "\" & Fil For Input As #f
         While Not EOF(f)
            Line Input #f, L
            If Not Left(L, 1) = "#" And Len(L) > 0 Then
               If Left(L, 5) = "[Sql-" Then
                  Ll = DonamParam(L)
                  If Left(L, 17) = "[Sql-LlistaCamps:" Then
                     CalCrearQuery = True
                     ReDim Camps(0)
                     While Len(Ll) > 0
                        ReDim Preserve Camps(UBound(Camps) + 1)
                        Camps(UBound(Camps)) = Car(Ll)
                     Wend
                  End If
               Else
                  Lll = L
                  ExecutaComandaSql "Delete paramsTPV Where Variable = 'Tarifa' and codiclient = " & llicencia
                  For i = 1 To UBound(Camps)
                     Select Case UCase(Camps(i))
                        Case "CODICLIENT":   Variable = "CodiBotiga"
                        Case "NOMMAQUINA":   Variable = "NpmMaquinaSw"
                        Case "TEXTE1":       Variable = "Capselera_1"
                        Case "TEXTE2":       Variable = "Capselera_2"
                        Case "TEXTE3":       Variable = "Capselera_3"
                        Case "TEXTE4":       Variable = "Capselera_4"
                        Case "TELEFON":      Variable = "Telefon"
                        Case "DIRECCIO":     Variable = "Direccio"
                        Case "NIF":          Variable = "Nif"
                        Case "SEMPRETICKET": Variable = "SempreTicket"
                        Case "LOGO":         Variable = "FileLogo"
                        Case Else:           Variable = Camps(i)
                     End Select
               
                     Valor = NormalitzaDes(Car(L))
                     If Not Variable = "Tarifa" Then
                        ExecutaComandaSql "Delete paramsTPV Where Variable = '" & Camps(i) & "' and codiclient = " & llicencia
                        ExecutaComandaSql "Delete paramsTPV Where Variable = '" & Variable & "' and codiclient = " & llicencia
                     End If
                     ExecutaComandaSql "insert into  paramsTPV (codiclient,Variable,Valor) Values (" & llicencia & ",'" & Variable & "','" & Valor & "') "
                  Next
               End If
            End If
            InformaEstat Estat, "", True
         Wend
         Close #f
      End If
      FitcherProcesat Fil
   Next
   
End Sub
Function BuscaTipus(Tip As Integer, s As String) As Variant
   
   Select Case Tip
      Case rdTypeCHAR:          BuscaTipus = s
      Case rdTypeGUID:          BuscaTipus = s
      Case rdTypeNUMERIC:       If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeDECIMAL:       If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeINTEGER:       BuscaTipus = 0
                                If IsNumeric(s) Then
                                    BuscaTipus = Val(s)
                                Else
                                   If s = "Falso" Then BuscaTipus = 0
                                   If s = "Verdadero" Then BuscaTipus = 1
                                End If
      Case rdTypeSMALLINT:      If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeFLOAT:         If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeREAL:          If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeDOUBLE:        If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeDATE:          If IsDate(s) Then BuscaTipus = CVDate(s)
      Case rdTypeTIME:          If IsDate(s) Then BuscaTipus = CVDate(s)
      Case rdTypeTIMESTAMP:     If IsDate(s) Then BuscaTipus = CVDate(s)
      Case rdTypeVARCHAR:       BuscaTipus = s
      Case rdTypeLONGVARCHAR:   If IsNumeric(s) Then BuscaTipus = Val(s)
      Case rdTypeBINARY:        BuscaTipus = s
      Case rdTypeVARBINARY:     BuscaTipus = s
      Case rdTypeLONGVARBINARY: BuscaTipus = s
      Case rdTypeBIGINT:        BuscaTipus = s
      Case rdTypeTINYINT:       BuscaTipus = s
      Case rdTypeBIT, -9:
         Select Case s
            Case "Falso": BuscaTipus = 0
            Case "Verdadero": BuscaTipus = 1
            Case "1": BuscaTipus = 1
            Case "0": BuscaTipus = 0
            Case Else: BuscaTipus = s
         End Select
      Case Else:
         BuscaTipus = s
   End Select
   
   If VarType(BuscaTipus) = Empty Then BuscaTipus = Null
   
End Function



Sub Interpreta_SqlTrans_Succeit(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double
   Dim FTPPATH As String, FTPUSER As String, FTPPASS As String, codiBotiga As String, File As String
   Dim IdCabeTPVLin, IdCabeTPV, Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset
   Dim Total As Double, TotalDesc As Double, TotalIva1 As Double, TotalIva2 As Double, TotalIva3 As Double, CosteIva As Double
   Dim CosteIva1 As Double, CosteIva2 As Double, CosteIva3 As Double, tipoIva As Integer, client, ClientAnt
   Dim BaseImpon As Double, CuotaIva, clientsCad As String
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*ComandaModificada*")
   ReDim FilsAImportar(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      Fil = Dir
   Wend
   
   If UBound(FilsAImportar) > 0 Then
      If ExisteixTaula("Servit_Temporal") Then EsborraTaula "Servit_Temporal"
      CreaTaulaServit "Servit_Temporal", False
      ExecutaComandaSql "ALTER TABLE Servit_Temporal ADD DiaDesti VarChar(255)  NULL"
   Else
      Exit Sub
   End If
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  On Error Resume Next
                  Db.Execute Ll
                  On Error GoTo 0
               End If
               If Left(L, 8) = "[Sql-Db:" Then
                  NomTaula = Ll
                  
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  sql = "Insert Into Servit_Temporal (DiaDesti," & LlistaC & ") Values (?," & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
               Q.rdoParameters(0) = NomTaula
               Lll = L
               For i = 0 To UBound(Camps) - 1
                  Valor = NormalitzaDes(Car(L))
On Error Resume Next
                  Q.rdoParameters(i + 1) = BuscaTipus(Q.rdoParameters(i + 1).Type, Valor)
On Error GoTo 0
               Next
               
'               If Camps(9) = "QuantitatDemanada" And Camps(10) = "QuantitatTornada" And Q.rdoParameters(9) = 0 And Q.rdoParameters(10) > 0 Then
'                    Q.rdoParameters(9) = Q.rdoParameters(10)
'                    Q.rdoParameters(10) = 0
'               End If
               
               ExecutaQuery Q
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
    'Insert TPVCabe y TPVLine
    If EmpresaActual = "Armengol" Then
        'Datos FTP
        FTPPATH = "localhost"
        FTPUSER = "armengol"
        FTPPASS = "fornarm"
        Set Rs2 = Db.OpenResultset("select distinct [diaDesti] as dataServit from [Servit_Temporal] ")
        Do While Not Rs2.EOF
            Set Rs3 = Db.OpenResultset("select distinct client from [Servit_Temporal] ")
            'clientsCad = "("
            Do While Not Rs3.EOF
                'clientsCad = clientsCad & "'" & Rs3("client") & "',"
                ExecutaComandaSql "insert into sincronitzaTangramDbg (tmst, p1, p2, p3, p4, p5) values (getdate(), '[" & Rs2("dataServit") & "]', '[" & Rs3("client") & "]', '[" & FTPPATH & "]', '[" & FTPUSER & "]', '[" & FTPPASS & "]')"
                
                InsertFeineaAFer "SincronitzaTangram", "[ " & Mid(Rs2("dataServit"), 9, 2) & "/" & Mid(Rs2("dataServit"), 6, 2) & "/" & Mid(Rs2("dataServit"), 1, 4) & "|('" & Rs3("client") & "')]", "[" & FTPPATH & "]", "[" & FTPUSER & "]", "[" & FTPPASS & "]", "[]"
                Rs3.MoveNext
                'If Rs3.EOF Then clientsCad = Mid(clientsCad, 1, (Len(clientsCad) - 1)) & ")"
            Loop
            'InsertFeineaAFer "SincronitzaTangram", "[ " & Mid(Rs2("dataServit"), 9, 2) & "/" & Mid(Rs2("dataServit"), 6, 2) & "/" & Mid(Rs2("dataServit"), 1, 4) & "|" & clientsCad & "]", "[" & FTPPATH & "]", "[" & FTPUSER & "]", "[" & FTPPASS & "]", "[]"
            Rs2.MoveNext
        Loop
    End If

   
   If ExisteixTaula("Servit_Temporal") Then
      InformaEstat Estat, "Actualitzant : Facturació "
      ExecutaComandaSql "Update Servit_Temporal Set Hora = 70,[timestamp] = getdate() "
      FiltraRegistresImportats Estat, "Servit_Temporal"
   End If
   
   EsborraTaula "Servit_Temporal"
   EsborraTaula "Servit_Temporal2"


End Sub



Sub Interpreta_SqlTrans_ComandaRevisada(Optional Estat As Label = Nothing)
   
    Dim CalCrearQuery As Boolean, FilsAImportar() As String
    Dim Fil As String, f, L As String, Ll As String, sql As String, Tipus() As String, Valor As String
    Dim LlistaCamps As String, NomTaula As String, botiga As String, DataComanda As String, Camps() As String
    Dim Kk As Integer, PerLine As Double, i As Integer, iD() As String
    Dim Q As rdoQuery

'Exit Sub
    CalCrearQuery = True
    Fil = Dir(AppPath & "\*ComandaRevisada*")
    ReDim FilsAImportar(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 12) = "[Sql-Botiga:" Then
                    botiga = Ll
               End If
               If Left(L, 8) = "[Sql-Db:" Then
                  NomTaula = Ll
               End If
               If Left(L, 9) = "[Sql-Dia:" Then
                  DataComanda = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                   sql = "Insert Into [" & NomTaulaRevisats(DateSerial(Split(DataComanda, "-")(0), Split(DataComanda, "-")(1), Split(DataComanda, "-")(2))) & "] "
                   sql = sql & "(Botiga, DataComanda, DataRevisio, Article, Viatge, Equip, Dependenta, Estat, Aux) Values "
                   sql = sql & "(?, ?, ?, ?, ?, ?, ?, ?, ?) "
                   Set Q = Db.CreateQuery("", sql)
On Error Resume Next
                  Q.rdoParameters(0) = botiga
                  Q.rdoParameters(1) = DataComanda
                  Q.rdoParameters(2) = BuscaTipus(Q.rdoParameters(2).Type, NormalitzaDes(Car(L))) 'Data de Revisió
                  iD = Split(NormalitzaDes(Car(L)), ",") 'Article, viatge, equip
                  Q.rdoParameters(3) = BuscaTipus(Q.rdoParameters(3).Type, iD(0)) 'Article
                  Q.rdoParameters(4) = BuscaTipus(Q.rdoParameters(4).Type, iD(1)) 'Viatge
                  Q.rdoParameters(5) = BuscaTipus(Q.rdoParameters(5).Type, iD(2)) 'Equip
                  Q.rdoParameters(6) = BuscaTipus(Q.rdoParameters(6).Type, NormalitzaDes(Car(L))) 'Dependenta
                  Q.rdoParameters(7) = BuscaTipus(Q.rdoParameters(7).Type, NormalitzaDes(Car(L))) 'Estat
                  Q.rdoParameters(8) = BuscaTipus(Q.rdoParameters(8).Type, NormalitzaDes(Car(L))) 'Aux
                             
                  ExecutaQuery Q
                  CalCrearQuery = False
               End If
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
End Sub




Sub Interpreta_SqlTrans_Teclat(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String, llicencia As Double
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double, FS, Fss, FilsAImportarData() As Date
   
   CalCrearQuery = True
   
   Set FS = CreateObject("Scripting.FileSystemObject")
   
   Fil = Dir(AppPath & "\*PreferenciasTeclat.SqlTrans")
   ReDim FilsAImportar(0)
   ReDim FilsAImportarData(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      ReDim Preserve FilsAImportarData(UBound(FilsAImportarData) + 1)
      Set Fss = FS.GetFile(AppPath & "\" & Fil)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      FilsAImportarData(UBound(FilsAImportarData)) = Fss.DateLastModified
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
      Dim p1, P2
      p1 = InStr(Fil, "TeclatToc_")
      llicencia = 0
      If p1 > 0 Then
         P2 = InStr(p1, Fil, "]")
         If P2 > 0 Then
            p1 = p1 + 10
            llicencia = 0
            If IsNumeric(Mid(Fil, p1, P2 - p1)) Then llicencia = Mid(Fil, p1, P2 - p1)
         End If
      End If
      
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  On Error Resume Next
                  Db.Execute Ll
                  On Error GoTo 0
               End If
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  If Not ExisteixTaula("TeclatsTpv") Then ExecutaComandaSql "CREATE TABLE [TeclatsTpv] ([Data] [datetime] NULL ,[Llicencia] [int] NULL,[Maquina] [float] Null,[Dependenta][float] Null,[Ambient][nvarchar] (255) Null ,[Article][float] Null,[Pos][float] Null,[Color][float] Null)"
                  sql = "Insert Into TeclatsTpv (Data,Llicencia," & LlistaC & ") Values (?,?," & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
               Q.rdoParameters(0) = FilsAImportarData(Kk)
               Q.rdoParameters(1) = llicencia
               Lll = L
               For i = 0 To UBound(Camps) - 1
                  Valor = NormalitzaDes(Car(L))
On Error Resume Next
                  Q.rdoParameters(i + 2) = BuscaTipus(Q.rdoParameters(i + 2).Type, Valor)
On Error GoTo 0
               Next
               ExecutaQuery Q
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
'   If ExisteixTaula("Servit_Temporal") Then
'      InformaEstat Estat, "Actualitzant : Facturació "
'      FiltraRegistresImportats Estat, "Servit_Temporal"
'   End If
'
'   EsborraTaula "Servit_Temporal"
'   EsborraTaula "Servit_Temporal2"


End Sub

Sub Interpreta_SqlTrans_Referencies(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String, llicencia As Double
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double, FS, Fss, FilsAImportarData() As Date
   
   CalCrearQuery = True
   ReDim Camps(0)
   Set FS = CreateObject("Scripting.FileSystemObject")
   Fil = Dir(AppPath & "\*CodisBarresReferencies.SqlTrans")
   ReDim FilsAImportar(0)
   ReDim FilsAImportarData(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      ReDim Preserve FilsAImportarData(UBound(FilsAImportarData) + 1)
      Set Fss = FS.GetFile(AppPath & "\" & Fil)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      FilsAImportarData(UBound(FilsAImportarData)) = Fss.DateLastModified
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : Referencies"
      Dim p1, P2
      
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  On Error Resume Next
                  Db.Execute Ll
                  On Error GoTo 0
               End If
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  If Not ExisteixTaula("CodisBarresReferencies") Then ExecutaComandaSql "CREATE TABLE [CodisBarresReferencies] ([Num][nvarchar] (255) Null ,[Tipus][nvarchar] (255) Null ,[Estat][nvarchar] (255) Null ,[Data] [datetime] NULL ,[TmSt] [datetime] NULL ,[Param1][nvarchar] (255) Null ,[Param2][nvarchar] (255) Null ,[Param3][nvarchar] (255) Null ,[Param4][nvarchar] (255) Null)"
                  sql = "Insert Into CodisBarresReferencies (" & LlistaC & ") Values (" & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
                  sql = "Delete CodisBarresReferencies Where Num = ? "
                  Set Q2 = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
               For i = 0 To UBound(Camps) - 1
                  Valor = NormalitzaDes(Car(L))
On Error Resume Next
                  Q.rdoParameters(i) = BuscaTipus(Q.rdoParameters(i).Type, Valor)
                  If i = 0 Then Q2.rdoParameters(i) = BuscaTipus(Q2.rdoParameters(i).Type, Valor)
On Error GoTo 0
               Next
               ExecutaQuery Q2
               ExecutaQuery Q
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
'   If ExisteixTaula("Servit_Temporal") Then
'      InformaEstat Estat, "Actualitzant : Facturació "
'      FiltraRegistresImportats Estat, "Servit_Temporal"
'   End If
'
'   EsborraTaula "Servit_Temporal"
'   EsborraTaula "Servit_Temporal2"


End Sub


Sub Interpreta_SqlTrans_Dedos(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String, llicencia As Double
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String, bytData() As Byte
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double, FS, Fss, FilsAImportarData() As Date
   
   CalCrearQuery = True
   ReDim Camps(0)
   Set FS = CreateObject("Scripting.FileSystemObject")
   Fil = Dir(AppPath & "\*Dedos.SqlTrans")
   ReDim FilsAImportar(0)
   ReDim FilsAImportarData(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      ReDim Preserve FilsAImportarData(UBound(FilsAImportarData) + 1)
      Set Fss = FS.GetFile(AppPath & "\" & Fil)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      FilsAImportarData(UBound(FilsAImportarData)) = Fss.DateLastModified
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : Referencies"
      Dim p1, P2
      
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  sql = "Insert Into " & DonamNomTaulaDedos() & " (" & LlistaC & ") Values (" & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
                  sql = "Delete " & DonamNomTaulaDedos() & "  Where usuario = ? "
                  Set Q2 = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
               For i = 0 To UBound(Camps) - 1
                  Valor = NormalitzaDes(Car(L))
On Error GoTo 0
'On Error Resume Next
                  If i = 1 Then
                    ReDim bytData(Len(Valor))
                    bytData = Valor
'                    Q.rdoParameters(i).AppendChunk(0) = BuscaTipus(Q.rdoParameters(i).Type, Valor)
                    Q.rdoParameters(i).Value = BuscaTipus(Q.rdoParameters(i).Type, Valor)
                  Else
                    Q.rdoParameters(i).Value = BuscaTipus(Q.rdoParameters(i).Type, Valor)
                  End If
                  
                  If i = 0 Then Q2.rdoParameters(i) = BuscaTipus(Q2.rdoParameters(i).Type, Valor)
'On Error GoTo 0
               Next
               ExecutaQuery Q2
               ExecutaQuery Q
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
'   If ExisteixTaula("Servit_Temporal") Then
'      InformaEstat Estat, "Actualitzant : Facturació "
'      FiltraRegistresImportats Estat, "Servit_Temporal"
'   End If
'
'   EsborraTaula "Servit_Temporal"
'   EsborraTaula "Servit_Temporal2"


End Sub



Sub Interpreta_SqlTrans_Comanda(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String, llicencia As Double
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double, FS, Fss, FilsAImportarData() As Date
   
   CalCrearQuery = True
   
   Set FS = CreateObject("Scripting.FileSystemObject")
   
   Fil = Dir(AppPath & "\*PreferenciasTeclat.SqlTrans")
   ReDim FilsAImportar(0)
   ReDim FilsAImportarData(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      ReDim Preserve FilsAImportarData(UBound(FilsAImportarData) + 1)
      Set Fss = FS.GetFile(AppPath & "\" & Fil)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      FilsAImportarData(UBound(FilsAImportarData)) = Fss.DateLastModified
      Fil = Dir
   Wend
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
      InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
      Dim p1, P2
      p1 = InStr(Fil, "TeclatToc_")
      llicencia = 0
      If p1 > 0 Then
         P2 = InStr(p1, Fil, "]")
         If P2 > 0 Then
            p1 = p1 + 10
            llicencia = Mid(Fil, p1, P2 - p1)
         End If
      End If
      
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  On Error Resume Next
                  Db.Execute Ll
                  On Error GoTo 0
               End If
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  For i = 1 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  If Not ExisteixTaula("TeclatsTpv") Then ExecutaComandaSql "CREATE TABLE [TeclatsTpv] ([Data] [datetime] NULL ,[Llicencia] [int] NULL,[Maquina] [float] Null,[Dependenta][float] Null,[Ambient][nvarchar] (255) Null ,[Article][float] Null,[Pos][float] Null,[Color][float] Null)"
                  sql = "Insert Into TeclatsTpv (Data,Llicencia," & LlistaC & ") Values (?,?," & LlistaCc & ") "
                  Set Q = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
               Q.rdoParameters(0) = FilsAImportarData(Kk)
               Q.rdoParameters(1) = llicencia
               Lll = L
               For i = 0 To UBound(Camps) - 1
                  Valor = NormalitzaDes(Car(L))
On Error Resume Next
                  Q.rdoParameters(i + 2) = BuscaTipus(Q.rdoParameters(i + 2).Type, Valor)
On Error GoTo 0
               Next
               ExecutaQuery Q
               PerLine = PerLine + 1
               If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
   If ExisteixTaula("Servit_Temporal") Then
      InformaEstat Estat, "Actualitzant : Facturació "
      FiltraRegistresImportats Estat, "Servit_Temporal"
   End If
   
   EsborraTaula "Servit_Temporal"
   EsborraTaula "Servit_Temporal2"


End Sub

Sub Interpreta_SqlTrans_Segurata(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, sql As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer, Contingut As String, NomMaquina As String, nomfile As String
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Rs As rdoResultset
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim codiBotiga As Double, Basura As String, Ddata As Date
   Dim O_CodiBotiga As Double, O_IdT As Double, O_Data As Date, CodiDependenta As Double
   
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\[Contingut#HoresSegurata]*.SqlTrans")
   ReDim FilsAImportar(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      Fil = Dir
   Wend
   
   If Not ExisteixTaula("HoresSegurata") Then
      ExecutaComandaSql "CREATE TABLE [HoresSegurata] ([Data] [datetime] NULL ,[IdT] [nvarchar] (255) NULL ,[QueFa] [char] (20) NULL ,[NomMaquina] [char] (20) NULL ,[CodiBotiga] [int] NULL,[TimeStamp] [datetime] NULL ,[Id] [uniqueidentifier] NULL ,[QuiStamp] [char] (10) NULL ) ON [PRIMARY]"
      sql = "ALTER TABLE [HoresSegurata] WITH NOCHECK ADD "
      sql = sql & "CONSTRAINT [DF_HoresSegurata_TimeStamp] DEFAULT (getdate()) FOR [TimeStamp], "
      sql = sql & "CONSTRAINT [DF_HoresSegurata_Id] DEFAULT (newid()) FOR [Id], "
      sql = sql & "CONSTRAINT [DF_HoresSegurata_QuiStamp] DEFAULT (host_name()) FOR [QuiStamp] "
      ExecutaComandaSql sql
      
      sql = "CREATE TRIGGER [DHoresSegurata] ON [HoresSegurata] "
      sql = sql & "FOR UPDATE AS "
      sql = sql & "Update [HoresSegurata] "
      sql = sql & "Set [TimeStamp] = GetDate(),"
      sql = sql & "    [QuiStamp]  = Host_Name() "
      sql = sql & "Where Id In (Select Id From Inserted) "
      ExecutaComandaSql sql
   End If
   
   If UBound(FilsAImportar) > 0 Then
      If ExisteixTaula("Temporal_HoresSegurata") Then EsborraTaula "Temporal_HoresSegurata"
      ExecutaComandaSql "CREATE TABLE [Temporal_HoresSegurata] ([Data] [datetime] NULL ,[IdT] [nvarchar] (255) NULL ,[QueFa] [char] (20) NULL ,[NomMaquina] [char] (20) NULL ,[CodiBotiga] [int] NULL ) ON [PRIMARY]"
      sql = "Insert Into Temporal_HoresSegurata ([Data],[IdT],[QueFa],[NomMaquina],[CodiBotiga]) Values (?,?,?,?,?) "
      Set Q = Db.CreateQuery("", sql)
   End If
   
   For Kk = 1 To UBound(FilsAImportar)
      Fil = FilsAImportar(Kk)
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      Q.rdoParameters(3) = NomMaquina
      f = FreeFile
      InformaEstat Estat, "Interpretant : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         My_DoEvents
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 17) = "[Sql-LlistaCamps:" Then DestriaCamps Ll, Camps, Tipus
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
                  Q.rdoParameters(0) = codiBotiga
                  If InStr(NomTaula, "_") > 0 Then
                     codiBotiga = Right(NomTaula, Len(NomTaula) - InStr(NomTaula, "_"))
                     Q.rdoParameters(4) = codiBotiga
                  End If
               End If
            Else
               For i = 1 To UBound(Camps)
                  Select Case UCase(Camps(i))
                     Case "DATA": Q.rdoParameters(0) = BuscaTipus(Q.rdoParameters(0).Type, NormalitzaDes(Car(L)))
                     Case "IDT": Q.rdoParameters(1) = BuscaTipus(Q.rdoParameters(1).Type, NormalitzaDes(Car(L)))
                     Case "QUEFA": Q.rdoParameters(2) = BuscaTipus(Q.rdoParameters(2).Type, NormalitzaDes(Car(L)))
                     Case Else: Basura = Car(L)
                  End Select
               Next
               If Len(Q.rdoParameters(1)) > 0 Then Q.Execute
            End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      InformaEstat Estat, "", True
   Next
   
   If UBound(FilsAImportar) > 0 Then
      InformaEstat Estat, "Actualitzant : Hores Arribada "
      If ExisteixTaula("Temporal_HoresSegurata") Then
         Set Rs = Db.OpenResultset("SELECT MAX(Data) As D2, MIN(Data) As D1 , NomMaquina From Temporal_HoresSegurata GROUP BY NomMaquina ")
         Set Q = Db.CreateQuery("", "Delete HoresSegurata Where NomMaquina = ? And Data >= ? And Data <= ?")
         While Not Rs.EOF
            Q.rdoParameters(0) = Trim(Rs("NomMaquina"))
            Q.rdoParameters(1) = Rs("D1")
            Q.rdoParameters(2) = Rs("D2")
            Q.Execute
            Rs.MoveNext
         Wend
         Rs.Close
         ExecutaComandaSql "INSERT INTO HoresSegurata (Data, IdT, QueFa, NomMaquina, CodiBotiga) SELECT Data, IdT, QueFa, NomMaquina, CodiBotiga From Temporal_HoresSegurata ORDER BY data DESC "
         ImportaDadesFichadorATpv "HoresSegurata"
         If Not ExisteixTaula("HoresSegurataNoImportades") Then ExecutaComandaSql "CREATE TABLE [HoresSegurataNoImportades] ([Data] [datetime] NULL ,[IdT] [nvarchar] (255) NULL ,[QueFa] [char] (20) NULL ,[NomMaquina] [char] (20) NULL ,[CodiBotiga] [int] NULL ) ON [PRIMARY]"
         ExecutaComandaSql "INSERT INTO HoresSegurataNoImportades ([Data], [IdT], [QueFa], [NomMaquina], [CodiBotiga]) SELECT [Data], [IdT], [QueFa], [NomMaquina], [CodiBotiga] FROM Temporal_HoresSegurata "
      End If
   End If
   EsborraTaula "Temporal_HoresSegurata"
   
End Sub



Sub ImportaDadesFichadorATpv(NomTaula As String)
   Dim Ddata As Date, Rs As rdoResultset, O_CodiBotiga As Double, O_IdT As String, O_Data As Date, CodiDependenta As Double, Q1 As rdoQuery, Q2 As rdoQuery
   
   Screen.MousePointer = 11
   Set Rs = Db.OpenResultset("SELECT Dependentes.codi, QueFa, CodiBotiga, Data FROM " & NomTaula & " JOIN Dependentes ON " & NomTaula & ".IdT = Dependentes.Tid ORDER BY Data")
         O_CodiBotiga = -1
         O_IdT = ""
         O_Data = DateAdd("y", 3, Now)
         
         While Not Rs.EOF
            CodiDependenta = Rs("Codi")
            Ddata = DateSerial(Year(Rs("Data")), Month(Rs("Data")), Day(Rs("Data")))
            O_CodiBotiga = Rs("CodiBotiga")
            If Not O_Data = Ddata Then
               CreaTaulesDadesTpv Ddata
               Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaHoraris(Ddata) & "] Where Botiga = ? And Data = ? ")
               Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaHoraris(Ddata) & "] (Botiga,Data,Dependenta,Operacio) Values (?,?,?,?)")
               O_Data = Ddata
            End If
                       
            Q1.rdoParameters(0) = O_CodiBotiga
            Q1.rdoParameters(1) = Rs("Data")
            Q1.Execute
            
            Q2.rdoParameters(0) = O_CodiBotiga
            Q2.rdoParameters(1) = Rs("Data")
            Q2.rdoParameters(2) = CodiDependenta
            If UCase(Trim(Rs("QueFa"))) = UCase("Plega") Then
               Q2.rdoParameters(3) = "P"
            Else
               Q2.rdoParameters(3) = "E"
            End If
            Q2.Execute
            My_DoEvents
         Rs.MoveNext
      Wend
      Rs.Close
   Screen.MousePointer = 0
   
End Sub


Sub Interpreta_SqlTrans_2(Patro As String, Optional Estat As Label = Nothing)
   Dim Fil As String, CalCrearQuery As Boolean, f, L As String, Ll As String, Lll As String, Files() As String, Iinici0 As Integer
   Dim Camps() As String, Tipus()  As String, NomTaula As String
   Dim LlistaCc As String, LlistaC As String, i As Integer, sql As String, Q As rdoQuery, CampsPerSet As String, LlistaCreate As String
   Dim FilesAImportar() As String, Kk As Integer, Q2 As rdoQuery
   Dim TeId As Boolean, SomDependentes As Boolean, SomDependentes_K As Integer, SomDependentes_P As Integer
   Dim ik As Integer, botiga As Double
   Dim TarifaCodi, Tarifanom
   Dim UltimIdInsertat As String
   
   Fil = Dir(AppPath & "\*" & Patro & ".SqlT*")
   ReDim Camps(0)
   ReDim Tipus(0)
   If Patro = "DeutesAnticips" Then If Not ExisteixTaula("DeutesAnticips") Then ExecutaComandaSql "CREATE TABLE DeutesAnticips ([Id]  [nvarchar] (255) Null ,Dependenta float NULL,Client  [nvarchar] (255) Null ,Data datetime NULL,Estat float NULL,Tipus float NULL,Import float NULL,Botiga float NULL,Detall [nvarchar] (255) Null )"

'On Error GoTo 0
   If UCase(Patro) = UCase("Dependentes") Then SomDependentes = True
   
   ReDim FilesAImportar(0)
   While Len(Fil) > 0
      ReDim Preserve FilesAImportar(UBound(FilesAImportar) + 1)
      FilesAImportar(UBound(FilesAImportar)) = Fil
      Fil = Dir
   Wend
   
   CalCrearQuery = True
   For Kk = 1 To UBound(FilesAImportar)
      f = FreeFile
      Fil = FilesAImportar(Kk)
      InformaEstat Estat, "Interpretant : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      TarifaCodi = ""
      Tarifanom = ""
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 13) = "[Sql-Execute:" Then
                  On Error Resume Next
'                  Db.Execute Ll
                  On Error GoTo 0
               End If
               If Left(L, 14) = "[Sql-NomTaula:" Then NomTaula = Ll
               If Left(L, 12) = "[Sql-Botiga:" Then botiga = Ll
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  ReDim Camps(0)
                  ReDim Tipus(0)
                  Lll = Ll
                  DestriaCamps Ll, Camps, Tipus
                  For i = 0 To UBound(Tipus)
                     If Tipus(i) = "Boolean" Then Tipus(i) = "Int"
                     If Tipus(i) = "Integer" Then Tipus(i) = "Int"
                     If Tipus(i) = "Byte" Then Tipus(i) = "bit"
                     If Tipus(i) = "TimeStamp" Then Tipus(i) = "smalldatetime"
                     If Tipus(i) = "Memo" Then Tipus(i) = "nvarchar"
                     If Tipus(i) = "Text" Then Tipus(i) = "nvarchar"
                     If SomDependentes And UCase(Camps(i)) = "TID" Then SomDependentes_K = i
                     If SomDependentes And UCase(Camps(i)) = "CODI" Then SomDependentes_P = i
                  Next
               End If
            Else
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
                  LlistaCreate = "("
                  TeId = False
                  For i = 1 To UBound(Camps)
                     If UCase(Camps(i)) = "ID" Then TeId = True
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     Select Case UCase(Camps(i))
                        Case "ID":
                                  If NomTaula = "ClientsFinalsPropietats" Or NomTaula = "Missatges" Or NomTaula = "ClientsFinals" Or NomTaula = "DeutesAnticips" Or NomTaula = "DeutesAnticipsV2" Then
                                      LlistaCreate = LlistaCreate & "[Id] [nvarchar] (255) "
                                  Else
                                     LlistaCreate = LlistaCreate & "[Id] [uniqueidentifier] Default (NEWID()    ) "
                                  End If
                        Case "TIMESTAMP": LlistaCreate = LlistaCreate & "[TimeStamp] [smalldatetime]    Default (GETDATE()  ) "
                        Case "QUISTAMP":  LlistaCreate = LlistaCreate & "[QuiStamp]  [nvarchar] (255)   Default (Host_Name()) "
                        Case "IMPORT": LlistaCreate = LlistaCreate & "[Import]  [Float] "
                        Case Else:
                           If Tipus(i) = "nvarchar" Then
                              LlistaCreate = LlistaCreate & "[" & Camps(i) & "] [" & Tipus(i) & "] (255) "
                           Else
                              LlistaCreate = LlistaCreate & "[" & Camps(i) & "] [" & Tipus(i) & "] "
                           End If
                     End Select
                     
                     If Not i = UBound(Camps) Then
                        LlistaCreate = LlistaCreate & ","
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  LlistaCreate = LlistaCreate & ")"
                  EsborraTaula NomTaula & "_Tmp "
                  ExecutaComandaSql "Create Table " & NomTaula & "_Tmp " & LlistaCreate
                  If Not ExisteixTaula(NomTaula) Then ExecutaComandaSql "Create Table " & NomTaula & " " & LlistaCreate
                  sql = "Insert Into " & NomTaula & "_Tmp (" & LlistaC & ") Values (" & LlistaCc & ") "
                  If NomTaula = "CertificatDeutesAnticips" Then
                    Set Q2 = Db.CreateQuery("", "update DeutesAnticips set estat = ? where id = ?")
                  End If
                  Set Q = Db.CreateQuery("", sql)
                  CalCrearQuery = False
               End If
On Error Resume Next
'On Error GoTo 0
               Lll = L

               If NomTaula = "CertificatDeutesAnticips" Then
                  Dim EsPagat As Boolean
                  EsPagat = False
                  For i = Iinici0 To UBound(Camps) - 1
                     Tarifanom = NormalitzaDes(Car(L))
                     If Camps(i + 1) = "IdDeute" Then Q.rdoParameters(1) = Tarifanom
                     If Camps(i + 1) = "Accio" Then If Tarifanom = "Pagat" Then EsPagat = True
'                     If Camps(i + 1) = "TmSt" Then TarifaCodi = Tarifanom
                  Next
                  Q.rdoParameters(0) = 0
                  If EsPagat Then
                    Q.rdoParameters(0) = 1
                    Q.Execute  ' De Moment sols certifikem els pagos
                  End If
                  L = Lll
               End If
                   
                   Iinici0 = 0
                   For i = Iinici0 To UBound(Camps) - 1
                      Q.rdoParameters(i) = BuscaTipus(Q.rdoParameters(i).Type, NormalitzaDes(Car(L)))
                      If TarifaCodi = "" And Camps(i + 1) = "TarifaCodi" Then TarifaCodi = Q.rdoParameters(i)
                      If Tarifanom = "" And Camps(i + 1) = "TarifaNom" Then Tarifanom = Q.rdoParameters(i)
                   Next
                  If SomDependentes Then Q.rdoParameters(SomDependentes_K - 1) = DependentaCodiTid(Q.rdoParameters(SomDependentes_P - 1))
                  If NomTaula = "CertificatDeutesAnticips" Then
                        Q.rdoParameters(8) = botiga
                        Q.rdoParameters(9) = Q.rdoParameters(0)
                        Q.rdoParameters(10) = botiga & "_" & Q.rdoParameters(0)
                  End If
                  
                  Q.Execute
On Error GoTo 0
            End If
         End If
         InformaEstat Estat, "", True
      Wend
      Close #f
      
      FitcherProcesat Fil
      
      If TeId Then
      
      
        If Patro = "ClientsFinalsPropietats" Then
            ExecutaComandaSql "delete " & NomTaula & " Where Id in (Select distinct id from [" & NomTaula & "_tmp] )"
            ExecutaComandaSql "Insert into [" & NomTaula & "] Select * from [" & NomTaula & "_tmp] )"
        End If
         If Patro = "DeutesAnticips" Or Patro = "DeutesAnticipsV2" Then
            ExecutaComandaSql "update deutesanticips set estat = 1 where botiga in(select distinct botiga from " & NomTaula & "_Tmp) and not id in(select id from " & NomTaula & "_Tmp)"
         End If
         GemeraSqlTrans NomTaula, Files
         If Not NomTaula = "ClientsFinals" Then If UBound(Files) > 0 Then frmSplash.IpConexio.AddMisatgeEnviar NomContingut(NomTaula), Files
         On Error Resume Next
         For ik = 1 To UBound(Files)
            Kill Files(ik)
         Next
         On Error GoTo 0
         
      Else
        If NomTaula = "TarifaEspecial" Then
           ExecutaComandaSql "Delete Tarifesespecials where tarifacodi = " & TarifaCodi
           ExecutaComandaSql "insert into tarifesespecials (tarifacodi,tarifanom,codi,preu,preumajor) select tarifacodi,tarifanom,codi,preu,preumajor from TarifaEspecial_tmp"
        ElseIf NomTaula = "CertificatDeutesAnticips" Then
           ExecutaComandaSql "Delete [" & NomTaula & "] where Params4 in  (Select Params4 From [" & NomTaula & "_Tmp]) "
           ExecutaComandaSql "Insert Into [" & NomTaula & "] Select * From [" & NomTaula & "_Tmp]"
        Else
           ExecutaComandaSql "Delete [" & NomTaula & "]"
           ExecutaComandaSql "Insert Into [" & NomTaula & "] Select * From [" & NomTaula & "_Tmp]"
        End If
      End If
   Next
   
End Sub




Sub BolcaAFitcher()
   
'   Set Q = Db.CreateQuery("", Sql)
'   Q.rdoParameters(0) = LastD
'   Set Rs = Q.OpenResultset
'
'   If Rs.EOF Then
'      Rs.Close
'      Exit Sub
'   End If
'
'   CampsClaus = ""
'   CampsCreate = ""
'   For i = 1 To Rs.rdoColumns.Count - 1
'      If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
'      CampsClaus = CampsClaus & "[" & Rs.rdoColumns(i).Name & "]" & "[" & DeTypeASt(Rs.rdoColumns(i).Type) & "]"
'   Next
'
'   QInsert.rdoParameters(1) = Null
'   QInsert.rdoParameters(2) = "[Sql-NomTaula:" & NomTaula & "]"
'   ExecutaQuery QInsert
'   QInsert.rdoParameters(2) = "[Sql-LlistaCamps:" & CampsClaus & "]"
'   ExecutaQuery QInsert
'
'   While Not Rs.EOF
'      Lin = ""
'      For i = 1 To Rs.rdoColumns.Count - 1
'         Lin = Lin & "[" & Normalitza(Rs.rdoColumns(i).Value) & "]"
'      Next
'      QInsert.rdoParameters(1) = Rs(0)
'      QInsert.rdoParameters(2) = Lin
'      ExecutaQuery QInsert
'      Rs.MoveNext
'   Wend
'   Rs.Close

End Sub

Sub CarregaPertinenses(CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String)
   Dim Rs As rdoResultset
   
   CondicioEnviamentClient = ""
   CondicioEnviamentViatge = ""
   CondicioEnviamentEquip = ""
   
   If Not ExisteixTaula("QueTinc") Then Exit Sub

   Set Rs = Db.OpenResultset("Select codi from clients ")
   CondicioEnviamentClient = ""
   While Not Rs.EOF
      If Not CondicioEnviamentClient = "" Then CondicioEnviamentClient = CondicioEnviamentClient & ","
      CondicioEnviamentClient = CondicioEnviamentClient & Rs(0)
      Rs.MoveNext
   Wend
   Rs.Close

   Set Rs = Db.OpenResultset("Select nom from viatges ")
   CondicioEnviamentViatge = ""
   While Not Rs.EOF
      If Not CondicioEnviamentViatge = "" Then CondicioEnviamentViatge = CondicioEnviamentViatge & ","
      CondicioEnviamentViatge = CondicioEnviamentViatge & "'" & Rs(0) & "'"
      Rs.MoveNext
   Wend
   Rs.Close

   Set Rs = Db.OpenResultset("Select nom from equipsdetreball  ")
   CondicioEnviamentEquip = ""
   While Not Rs.EOF
      If Not CondicioEnviamentEquip = "" Then CondicioEnviamentEquip = CondicioEnviamentEquip & ","
      CondicioEnviamentEquip = CondicioEnviamentEquip & "'" & Rs(0) & "'"
      Rs.MoveNext
   Wend
   Rs.Close

End Sub

Public Sub AccesAlRegistre(Clau As String, alor As String, Defecte As String, Optional EscriuDefecte As Boolean = False)
   Dim Resultat As String
   
   If EscriuDefecte Then SaveSetting "Hit", App.EXEName, Clau, Defecte
   
   Resultat = GetSetting("Hit", App.EXEName, Clau)
   
   If Len(Resultat) = 0 Then
      Resultat = Defecte
      SaveSetting "Hit", App.EXEName, Clau, Resultat
   End If
   
   alor = Resultat
   
End Sub


Sub EnviaFacturacioDia(TaulaDesti As String, NomTaula As String, LCamps As String, CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String, Optional CondicioExtra As String = "")
   Dim Rs As rdoResultset, Q As rdoQuery, NewD As Double, lin As String, c, StDia As String, sql As String, CampsClaus As String, CampsCreate As String, i As Integer
   Dim Ce As String
   
   If Not ExisteixTaula(NomTaula) Then Exit Sub
   
   sql = "Insert " & TaulaDesti & " (" & LCamps & ",[DiaDesti]) "
   sql = sql & "Select " & LCamps & ",'[" & NomTaula & "]' "
   sql = sql & "From  [" & NomTaula & "] "
   sql = sql & "Where " & CondicioExtra & " "
   
'   If Not TePermisPer("EnviaTotAFora") Then
'      If Len(CondicioEnviamentClient) > 0 And Len(CondicioEnviamentViatge) > 0 And Len(CondicioEnviamentViatge) > 0 Then
'         sql = sql & " And ("
'         sql = sql & "Client Not In(" & CondicioEnviamentClient & ") "
'         sql = sql & "Or Viatge Not In(" & CondicioEnviamentViatge & ") "
'         sql = sql & "Or Equip  Not In(" & CondicioEnviamentEquip & " ) "
'         sql = sql & "     )"
'      End If
'   End If
   
   ExecutaComandaSql sql
   
End Sub


Sub EnviaFacturacio(Files() As String)
'   Dim CampsClaus As String, CampsCreate As String, Lin As String, rs As rdoResultset, F, i As Integer, NomFileFeinaFeta As String, Condicio As String
'   Dim D As Date, Df As Date, CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String, NomFile As String, res As rdoResultset
'   Dim LastNomTaula As String, Nomtaula As String, Files1() As String, Lineas As Integer
'
'   CarregaPertinenses CondicioEnviamentClient, CondicioEnviamentViatge, CondicioEnviamentEquip
'
'   EsborraTaula "TemporalEnviaments"
'   CreaTaulaServit "TemporalEnviaments", False
'   ExecutaComandaSql "ALTER TABLE TemporalEnviaments ADD DiaDesti VarChar(255)  NULL"
'
'   ExecutaComandaSql "Delete Records Where Concepte = 'FacturacioNovaData' "
'   ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values('FacturacioNovaData',GetDate())"
'   Set res = Db.OpenResultset("Select * From Records Where Concepte = 'Facturacio'")
'   If res.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('Facturacio',DATEADD(Day, -1, GetDate()))"
'   res.Close
'
'   Set rs = Db.OpenResultset("Select * From TemporalEnviaments ")
'   CampsClaus = ""
'   CampsCreate = ""
'   For i = 0 To rs.rdoColumns.Count - 1
'      If Not rs.rdoColumns(i).Name = "DiaDesti" Then
'         If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
'         CampsClaus = CampsClaus & "[" & rs.rdoColumns(i).Name & "]" & "[" & DeTypeASt(rs.rdoColumns(i).Type) & "]"
'         CampsCreate = CampsCreate & "[" & rs.rdoColumns(i).Name & "]"
'      End If
'   Next
'
'   If ExisteixTaula("ComandesModificades") Then
'      Set rs = Db.OpenResultset("Select Distinct TaulaOrigen From ComandesModificades ")
'      While Not rs.EOF
'         My_DoEvents
'         Condicio = "Id In (Select Id From ComandesModificades Where  "
'         Condicio = Condicio & "   TaulaOrigen = '" & rs("TaulaOrigen") & "'  "
'         Condicio = Condicio & "   And [TimeStamp] >  (Select Max(TimeStamp) From  Records Where Concepte = 'Facturacio')  "
'         Condicio = Condicio & "   And [TimeStamp] <= (Select Max(TimeStamp) From  Records Where Concepte = 'FacturacioNovaData') "
'         Condicio = Condicio & "            ) "
'
'         EnviaFacturacioDia "TemporalEnviaments", rs("TaulaOrigen"), CampsCreate, CondicioEnviamentClient, CondicioEnviamentViatge, CondicioEnviamentEquip, Condicio
'         rs.MoveNext
'      Wend
'      rs.Close
'   End If
'   ExecutaComandaSql "Delete ComandesModificades Where  [TimeStamp] <= (Select Max(TimeStamp) From  Records Where Concepte = 'FacturacioNovaData') "
'
'   ReDim Files(0)
'   F = FreeFile
'   Lineas = 0
'   Set rs = Db.OpenResultset("Select * From TemporalEnviaments DiaDesti ", rdConcurRowVer)
'   If Not rs.EOF Then
'      LastNomTaula = ""
'      While Not rs.EOF
'         If Lineas <= 0 Then
'            EnviaFacturacioCreaFile Files, F, CampsClaus, "_SucceitFins.SqlTrans"
'            LastNomTaula = ""
'            Lineas = 500
'         End If
'
'         Lin = ""
'         For i = 0 To rs.rdoColumns.Count - 1
'            If rs.rdoColumns(i).Name = "DiaDesti" Then
'               Nomtaula = rs.rdoColumns(i).Value
'            Else
'               If rs.rdoColumns(i).Type = rdTypeDOUBLE Or rs.rdoColumns(i).Type = rdTypeDECIMAL Or rs.rdoColumns(i).Type = rdTypeFLOAT Or rs.rdoColumns(i).Type = rdTypeINTEGER Or rs.rdoColumns(i).Type = rdTypeNUMERIC Or rs.rdoColumns(i).Type = rdTypeREAL Or rs.rdoColumns(i).Type = rdTypeSMALLINT Then
'                  Lin = Lin & "[" & NormalitzaNumero(rs.rdoColumns(i).Value) & "]"
'               Else
'                  Lin = Lin & "[" & Normalitza(rs.rdoColumns(i).Value) & "]"
'               End If
'            End If
'         Next
'
'         If Not LastNomTaula = Nomtaula Then
'            Print #F, "[Sql-NomTaula:" & Nomtaula & "]"
'            LastNomTaula = Nomtaula
'         End If
'
'         Print #F, Lin
'         rs.MoveNext
'         Lineas = Lineas - 1
'      Wend
'      rs.Close
'      Close F
'   End If
'
'If UCase(EmpresaActual) = UCase("iblatpa") Then
'   FileCopy Files, "D:\Ftp\integraciones" & Files
'
'End If
'
'   EsborraTaula "TemporalEnviaments"
'   ExecutaComandaSql "Delete Records Where Concepte = 'Facturacio'"
'   ExecutaComandaSql "Update Records Set Concepte = 'Facturacio' Where Concepte = 'FacturacioNovaData' "

End Sub





Sub EnviaComandes(Files() As String, Param As String)
   Dim CampsClaus As String, CampsCreate As String, lin As String, Rs As rdoResultset, f, i As Integer, NomFileFeinaFeta As String
   Dim D As Date, Df As Date, CondicioEnviamentClient As String, CondicioClient As String, nomfile As String, res As rdoResultset
   Dim LastNomTaula As String, NomTaula As String, Files1() As String, Lineas As Integer, Condicio As String, CalBorrar As String
   Dim Rs2 As rdoResultset
   
On Error GoTo Fi

    If UCase(EmpresaActual) = UCase("Tena") Then
        Set Rs = Db.OpenResultset("Select top 1 * From Servit_Tmp Where (left(Viatge,1)='.' or left(Equip,1)='.') And Client = " & Param & " ")
    Else
        Set Rs = Db.OpenResultset("Select top 1 * From Servit_Tmp Where Client = " & Param & " ")
    End If

   CampsClaus = ""
   CampsCreate = ""
   For i = 0 To Rs.rdoColumns.Count - 1
      If Not Rs.rdoColumns(i).Name = "DiaDesti" Then
         If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
         CampsClaus = CampsClaus & "[" & Rs.rdoColumns(i).Name & "]" & "[" & DeTypeASt(Rs.rdoColumns(i).Type) & "]"
         CampsCreate = CampsCreate & "[" & Rs.rdoColumns(i).Name & "]"
      End If
   Next

   f = FreeFile
   Lineas = 0
   If UCase(EmpresaActual) = UCase("Tena") Then
        Set Rs = Db.OpenResultset("Select * From Servit_Tmp Where (left(Viatge,1)='.' or left(Equip,1)='.') And Client = " & Param & " Order By DiaDesti", rdConcurRowVer)
   Else
        '11/02/2011 JORGE
        'Modificado para enviar los 0 de los productos que tienen marcado fixat a graella
        'Set rs = Db.OpenResultset("Select * From Servit_Tmp Where (QuantitatDemanada<>0 or QuantitatServida <>0) And  Client = " & Param & " Order By DiaDesti", rdConcurRowVer)
        '********************
        Set Rs = Db.OpenResultset("Select * From Servit_Tmp Where ((QuantitatDemanada<>0 or QuantitatServida <>0) or CodiArticle in (select CodiArticle from ArticlesPropietats where Variable = 'FixaGraella' and Valor = 'on')) And  Client = " & Param & " Order By DiaDesti", rdConcurRowVer)
        '********************
   End If
      
   If Not Rs.EOF Then
      LastNomTaula = ""
      NomTaula = ""
      While Not Rs.EOF
         lin = ""
         For i = 0 To Rs.rdoColumns.Count - 1
            If Rs.rdoColumns(i).Name = "DiaDesti" Then
               NomTaula = Rs.rdoColumns(i).Value
            Else
               If Rs.rdoColumns(i).Type = rdTypeDOUBLE Or Rs.rdoColumns(i).Type = rdTypeDECIMAL Or Rs.rdoColumns(i).Type = rdTypeFLOAT Or Rs.rdoColumns(i).Type = rdTypeINTEGER Or Rs.rdoColumns(i).Type = rdTypeNUMERIC Or Rs.rdoColumns(i).Type = rdTypeREAL Or Rs.rdoColumns(i).Type = rdTypeSMALLINT Then
                  lin = lin & "[" & NormalitzaNumero(Rs.rdoColumns(i).Value) & "]"
               Else
                  lin = lin & "[" & Normalitza(Rs.rdoColumns(i).Value) & "]"
               End If
            End If
         Next

         If Not LastNomTaula = NomTaula Or Lineas <= 0 Then
            nomfile = EnviaFacturacioCreaFile(Files, f, CampsClaus, "_Comandes_" & Param & ".SqlTrans")
            LastNomTaula = ""
            Lineas = 500
         End If
         
         If Not LastNomTaula = NomTaula Then
            Print #f, "[Sql-NomTaula:" & NomTaula & "]"
            Print #f, "[Sql-Data:" & Left(Right(NomTaula, 9), 8) & "]"
            Print #f, "[Sql-Client:" & Param & "]"
            Set Rs2 = Db.OpenResultset("Select count(*)  from [" & DonamNomTaulaServit(DateSerial(Split(NomTaula, "-")(0), Split(NomTaula, "-")(1), Split(NomTaula, "-")(2))) & "] Where Client = " & Param & "  and not tipuscomanda = 2 and not id in (Select id From Servit_Tmp Where Client = " & Param & " )")
            
'            CalBorrar = "Delete * From Servit " ' Where Client = " & Param & " "
'            If Not Rs2.EOF Then If Not IsNull(Rs2(0)) Then If Rs2(0) > 0 Then CalBorrar = ""
'            If Not CalBorrar = "" Then Print #f, "[Sql-Execute:" & CalBorrar & "]"
            Print #f, "[Sql-Execute: Delete * From Servit ]"
            Rs2.Close
            LastNomTaula = NomTaula
         End If

         Print #f, lin
         Rs.MoveNext
         Lineas = Lineas - 1
      Wend
      Rs.Close
            
      Close f
        
   End If

Fi:
End Sub
Sub GemeraConfiguracio(llicencia As String, Files() As String)
   Dim Rs As rdoResultset, Q As rdoQuery, f, CampsCreate As String, c, CampsClaus As String, lin As String, Codi As Double, CodiClient As Double, Vaa As String
   
   ReDim Files(0)
   CodiClient = LlicenciaCodiClient(llicencia)
   'If CodiClient = -1 Then CodiClient = Llicencia
   
   CampsCreate = ""
   lin = ""
   
   Set Rs = Db.OpenResultset("Select Nom,Codi As CodiBotiga,Codi As CodiClient,[Preu Base],Nif,Adresa,Ciutat,Cp,Lliure,[Nom Llarg],[Tipus Iva],[Desconte 5]  From Clients Where Codi = " & CodiClient & " ")
   If Not Rs.EOF And Not Codi = -1 Then
      For Each c In Rs.rdoColumns
         CampsCreate = CampsCreate & "[" & Normalitza(c.Name) & "]"
      Next
      
      For Each c In Rs.rdoColumns
         Vaa = Normalitza(c.Value)
         If UCase(c.Name) = UCase("Preu Base") Then
            If CodiClientTeTpv(CodiClient) Then Vaa = "1"
         End If
         lin = lin & "[" & Vaa & "]"
      Next
   End If
   Rs.Close
   
   Set Q = Db.CreateQuery("", "Select * From ParamsTpv Where CodiClient = ? order by variable ")
   Q.rdoParameters(0) = CodiClient
   Set Rs = Q.OpenResultset()
   
   If Not Rs.EOF Then
      While Not Rs.EOF
         If InStr(CampsCreate, "[" & Rs("Variable") & "]") = 0 Or Rs("Variable") = "Tarifa" Then
            CampsCreate = CampsCreate & "[" & Normalitza(Rs("Variable")) & "]"
            lin = lin & "[" & Normalitza(Rs("Valor")) & "]"
         End If
         Rs.MoveNext
      Wend
      Rs.Close
   End If
  
   If Len(lin) > 0 Or Len(CampsCreate) > 0 Then
      ReDim Files(1)
      Files(1) = AppPath & "\ConfiguracioTpv.Cfg"
      On Error Resume Next
         FitcherProcesat Files(1), True, True
      On Error GoTo 0
      f = FreeFile
      Open Files(1) For Output As #f
      Print #f, "[Sql-LlistaCamps:" & CampsCreate & "]"
      Print #f, lin
      Close f
   End If
End Sub


Sub GemeraSqlTrans(NomTaula As String, Files() As String, Optional sql2 As String, Optional DeteteSi As Boolean = False)
    Dim f, res As rdoResultset, lin As String, c, CampsClaus As String, CampsCreate As String, sql As String
    Dim TeTimeStamp As Boolean, Chunk As Boolean, Camp As String
    
    ReDim Files(0)
    
    On Error GoTo NoTaula
    
    If Len(sql2) > 0 Then
        Set res = Db.OpenResultset(sql2)
    Else
        Set res = Db.OpenResultset("Select * From " & NomTaula & " ")
    End If
    
    If Not res.EOF Then
        TeTimeStamp = False
        Chunk = False
        CampsClaus = ""
        CampsCreate = ""
        Camp = "TimeStamp"
        If NomTaula = "DeutesAnticips" Then Camp = "Data"
        If NomTaula = "DeutesAnticipsV2" Then Camp = "Data"
        
        For Each c In res.rdoColumns
            If UCase(c.Name) = UCase(Camp) Then TeTimeStamp = True
            If c.ChunkRequired Then Chunk = True
            If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
            CampsClaus = CampsClaus & "[" & c.Name & "]" & "[" & DeTypeASt(c.Type) & "]"
            CampsCreate = CampsCreate & "[" & c.Name & "]" & " " & DeTypeASt(c.Type) & " "
        Next
        
        If TeTimeStamp Then
            Set res = Db.OpenResultset("Select * From Records Where Concepte = '" & NomTaula & "'")
            If res.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('" & NomTaula & "',DATEADD(Day, -1, GetDate()))"
            res.Close
            
            sql = "Select * "
            sql = sql & "From " & NomTaula & " "
            sql = sql & "Where [" & Camp & "] >  (Select Max([TimeStamp]) From Records Where Concepte = '" & NomTaula & "') "
            If Len(sql2) > 0 Then sql = sql2
            Set res = Db.OpenResultset(sql, rdConcurRowVer)
        Else
            sql = "Select * From " & NomTaula & " "
            If Len(sql2) > 0 Then sql = sql2
        End If
        
        If Chunk Then Set res = Db.OpenResultset(sql, rdConcurRowVer) Else Set res = Db.OpenResultset(sql)
        
        If Not res.EOF Then
            On Error Resume Next
            ReDim Files(1)
            Files(1) = AppPath & "\" & NomTaula & ".SqlTrans"
            If Files(1) = AppPath & "\DeutesAnticipsv2.SqlTrans" Then Files(1) = AppPath & "\DeutesAnticips.SqlTrans"
            f = FreeFile
            FitcherProcesat Files(1), True, True
            
            Open Files(1) For Append As f
            Print #f, "[Sql-NomTaula:" & NomTaula & "]"
            If DeteteSi Then
                Print #f, "[Sql-Execute:Delete * From " & NomTaula & "]"
            Else
                Print #f, "[Sql-Execute:Delete " & NomTaula & "]"
            End If
        
            Print #f, "[Sql-LlistaCamps:" & CampsClaus & "]"
            While Not res.EOF
                lin = ""
                For Each c In res.rdoColumns
                    If Chunk Then
                        lin = lin & "[" & Normalitza(DameValor(res, c.Name)) & "]"
                    Else
                        lin = lin & "[" & Normalitza(c.Value) & "]"
                    End If
                Next
                
                If NomTaula = "ConstantsClient" Then
                    If InStr(lin, "NoPagaEnTienda") > 0 Then
                        While InStr(lin, "NoPagaEnTienda") > 0
                            lin = Left(lin, InStr(lin, "NoPagaEnTienda") - 1) & "NoPagaABotiga" & Right(lin, Len(lin) - InStr(lin, "NoPagaEnTienda") - 13)
                        Wend
                    End If
                End If
                Print #f, lin
                res.MoveNext
            Wend
            res.Close
            Close f
        End If
    End If
    
    ActualitzaTaula NomTaula   ' Insertem Ens Nous
    ExecutaComandaSql "Delete Records Where Concepte = '" & NomTaula & "'"
    ExecutaComandaSql "Insert Records (Concepte,[TimeStamp]) Select '" & NomTaula & "',(Select Max([" & Camp & "]) From " & NomTaula & " )"

NoTaula:

End Sub


Sub GemeraSqlTransBin(NomTaula As String, Files() As String, Optional sql2 As String)
   Dim f, res As rdoResultset, lin As String, c, CampsClaus As String, CampsCreate As String, sql As String
   Dim TeTimeStamp As Boolean, Chunk As Boolean, Camp As String
   
   On Error GoTo NoTaula
'   If Not ExisteixTaula(NomTaula) Then Exit Sub
   If Len(sql2) > 0 Then
      Set res = Db.OpenResultset(sql2)
   Else
      Set res = Db.OpenResultset("Select * From " & NomTaula & " ")
   End If
   
   If Not res.EOF Then
      TeTimeStamp = False
      Chunk = False
      CampsClaus = ""
      CampsCreate = ""
      Camp = "TimeStamp"
      If NomTaula = "DeutesAnticips" Then Camp = "Data"
      If NomTaula = "DeutesAnticipsV2" Then Camp = "Data"

      For Each c In res.rdoColumns
         If UCase(c.Name) = UCase(Camp) Then TeTimeStamp = True
         If c.ChunkRequired Then Chunk = True
         If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
         CampsClaus = CampsClaus & "[" & c.Name & "]" & "[" & DeTypeASt(c.Type) & "]"
         CampsCreate = CampsCreate & "[" & c.Name & "]" & " " & DeTypeASt(c.Type) & " "
      Next
      If TeTimeStamp Then
         Set res = Db.OpenResultset("Select * From Records Where Concepte = '" & NomTaula & "'")
         If res.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('" & NomTaula & "',DATEADD(Day, -1, GetDate()))"
         res.Close
         sql = "Select * "
         sql = sql & "From " & NomTaula & " "
         sql = sql & "Where [" & Camp & "] >  (Select Max([TimeStamp]) From Records Where Concepte = '" & NomTaula & "') "
If Len(sql2) > 0 Then sql = sql2
         Set res = Db.OpenResultset(sql, rdConcurRowVer)
      Else
         sql = "Select * From " & NomTaula & " "
         If Len(sql2) > 0 Then sql = sql2
      End If
      
      If Chunk Then Set res = Db.OpenResultset(sql, rdConcurRowVer) Else Set res = Db.OpenResultset(sql)
      If Not res.EOF Then
         On Error Resume Next
         ReDim Files(1)
         Files(1) = AppPath & "\" & NomTaula & ".SqlTrans"
         f = FreeFile
         FitcherProcesat Files(1), True, True
         
         Open Files(1) For Append As f
         Print #f, "[Sql-NomTaula:" & NomTaula & "]"
         Print #f, "[Sql-Execute:Delete " & NomTaula & "]"
         Print #f, "[Sql-LlistaCamps:" & CampsClaus & "]"
         While Not res.EOF
            lin = ""
            For Each c In res.rdoColumns
               If Chunk Then
                  lin = lin & "[" & Normalitza(DameValor(res, c.Name)) & "]"
               Else
                  lin = lin & "[" & Normalitza(c.Value) & "]"
               End If
            Next
            Print #f, lin
            res.MoveNext
         Wend
         res.Close
         Close f
      End If
   End If
   
   ActualitzaTaula NomTaula   ' Insertem Ens Nous
   ExecutaComandaSql "Delete Records Where Concepte = '" & NomTaula & "'"
   ExecutaComandaSql "Insert Records (Concepte,[TimeStamp]) Select '" & NomTaula & "',(Select Max([" & Camp & "]) From " & NomTaula & " )"

NoTaula:

End Sub

Sub GemeraSqlTransFiles(nom As String, Param As String, Files() As String, Contingut As String)
   Dim NomDistribucio As String, Par As String, Passatgers() As String, Fr, Fw, i As Integer, K As Integer, SubTamany As Double, crc As Double, kCrc As Double, NumLastPassatger As Double, NomEscriptor As String, Tam As Double
   Dim Buff() As Byte, AccTam As Double
   
   ReDim Passatgers(0)
   
   NomDistribucio = Car(Param)
   Contingut = NomDistribucio
   Par = Car(Param)
   While Len(Par) > 0
      ReDim Preserve Passatgers(UBound(Passatgers) + 1)
      Passatgers(UBound(Passatgers)) = Par
      Par = Car(Param)
   Wend
   
   For i = 1 To UBound(Passatgers)
      Fr = FreeFile
      InformaMiss "Preparant Passatger " & Passatgers(i), True
      Open AppPath & "\" & Passatgers(i) For Binary Access Read As Fr
      SubTamany = 0
      While SubTamany < LOF(Fr)
         ReDim Preserve Files(UBound(Files) + 1)
         Files(UBound(Files)) = "TmpDistribucio_" & i
         FitcherProcesat Files(UBound(Files))
         Fw = FreeFile
         Open AppPath & "\" & Files(UBound(Files)) For Binary Access Write As Fw
         Tam = 300000
         If (SubTamany + Tam) > LOF(Fr) Then Tam = LOF(Fr) - SubTamany
         Buff = InputB(Tam, Fr)
         SubTamany = SubTamany + UBound(Buff) + 1
         Put #Fw, 1, Buff
         Close Fw
         crc = 0
         For kCrc = 1 To UBound(Buff)
            crc = (crc + Asc(Buff(kCrc))) Mod 10000
         Next
         NomEscriptor = "[Bloc#" & Format(UBound(Files), "0000") & "]" & "[File#" & Passatgers(i) & "]" & "[Crc#" & crc & "]"
         FitcherProcesat NomEscriptor
         Name AppPath & "\" & Files(UBound(Files)) As AppPath & "\" & NomEscriptor
         Files(UBound(Files)) = NomEscriptor
      Wend
      Close Fr
   Next
   
   For i = 1 To UBound(Files)
      FitcherProcesat Files(i) & "[De#" & UBound(Files) & "].SqlTrans"
      Name AppPath & "\" & Files(i) As AppPath & "\" & Files(i) & "[De#" & UBound(Files) & "].SqlTrans"
      Files(i) = AppPath & "\" & Files(i) & "[De#" & UBound(Files) & "].SqlTrans"
   Next
   
End Sub

Sub GemeraSqlTransTpvVellsCodis(Files() As String)
   Dim f, Rs As rdoResultset, Preu As Double, NoAcabat As Boolean, IncPreu1 As Double, IncPct1 As Double, IncPreu2 As Double, IncPct2 As Double
   
   ReDim Files(0)
   
   On Error GoTo NoTaula
   If Not ExisteixTaula("Memotecnics") Then Exit Sub
   If Not ExisteixTaula("Articles") Then Exit Sub
   If Not ExisteixTaula("Atributs") Then Exit Sub
   
   On Error Resume Next
   ReDim Files(1)
   Files(1) = AppPath & "\TpvVellsCodis.SqlTrans"
   f = FreeFile
   FitcherProcesat Files(1), True, True
   
   Open Files(1) For Append As f
   Print #f, "[Sql-NomTaula:TpvVellsCodis]"
   Print #f, "[Sql-LlistaCamps:[CodiGenetic][Nom][Preu]]"
   
   Set Rs = Db.OpenResultset("Select CodiGenetic,Nom,Preu From Articles Order By CodiGenetic ")
   If Not Rs.EOF Then
      While Not Rs.EOF
         Print #f, "[" & Normalitza(Rs("CodiGenetic")) & "]" & "[" & Left(Normalitza(Rs("Nom")) & Space(20), 20) & "]" & "[" & Normalitza(Rs("Preu")) & "]"
         Rs.MoveNext
      Wend
      Rs.Close
   End If
   Print #f, "Eoeo"
   Set Rs = Db.OpenResultset("Select Articles.nom,Articles.Preu,ISNUMERIC(Memotecnic) As N,Memotecnic,Memotecnics.Codi,Atribut From Memotecnics  Join Articles On Memotecnics.Codi=Articles.Codi Where not Memotecnic in (Select Cast(CodiGenetic As nvarchar) From Articles) And Len(Memotecnic) < 5 Order By N Desc")
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         If Rs("N") = 0 Then Exit Do
         Preu = 0
         
         Preu = Rs("Preu")
         If Not Rs("Atribut") = 0 Then
            AtributCodiPreus Rs("Atribut"), NoAcabat, IncPreu1, IncPct1, IncPreu2, IncPct2
            Preu = Preu + IncPreu1
            Preu = Preu + (Preu * (IncPct1 / 100))
         End If
         
         Print #f, "[" & Normalitza(Trim(Rs("Memotecnic"))) & "]" & "[" & Left(Normalitza(Rs("Nom") & AtributCodiNomCurt(Rs("Atribut"))) & Space(20), 20) & "]" & "[" & Normalitza(Preu) & "]"
         Rs.MoveNext
      Loop

      Rs.Close
   End If
   
   Close f

NoTaula:

End Sub
Function AtributCodiNomCurt(Codi As Long) As String
   Dim sql As String, Codis_Dels_Atributs As rdoResultset
   
   AtributCodiNomCurt = ""
   
   sql = ("select TexteAnexat from atributs where Codi = " & Codi & " ")
   Set Codis_Dels_Atributs = Db.OpenResultset(sql)
   
   If Not Codis_Dels_Atributs.EOF Then
      AtributCodiNomCurt = Codis_Dels_Atributs("TexteAnexat")
   End If

End Function




Function DameValor(Rs As rdoResultset, NomCamp As String) As Variant
   
   If Rs(NomCamp).ChunkRequired Then
      
      If Rs(NomCamp).ColumnSize = 0 Then
         DameValor = ""
      Else
         DameValor = Rs(NomCamp).GetChunk(Rs(NomCamp).ColumnSize)
      End If
      
   Else
      DameValor = Rs(NomCamp)
   End If
   
End Function




Function DeTypeASt(n As Integer) As String
'rdTypeREAL
'rdTypeVARBINARY
'rdTypeLONGVARBINARY
'rdTypeTINYINT
'rdTypeSMALLINT
   Select Case n
      Case rdTypeBIGINT: DeTypeASt = "Big Integer"
      Case rdTypeBINARY: DeTypeASt = "Binary"
'      Case dbBoolean: DeTypeASt = "Boolean"
      Case rdTypeBIT: DeTypeASt = "Byte"
      Case rdTypeCHAR: DeTypeASt = "Char"
'      Case dbCurrency: DeTypeASt = "Currency"
      Case rdTypeTIME: DeTypeASt = "DateTime"
      Case rdTypeDECIMAL: DeTypeASt = "Decimal"
      Case rdTypeDOUBLE: DeTypeASt = "Double"
'      Case dbFloat: DeTypeASt = "Float"
'      Case dbGUID: DeTypeASt = "Guid"
      Case rdTypeINTEGER: DeTypeASt = "Integer"
      Case rdTypeFLOAT: DeTypeASt = "float"
'      Case dbLongBinary: DeTypeASt = "float Binary"
      Case rdTypeLONGVARCHAR: DeTypeASt = "Memo"
      Case rdTypeNUMERIC: DeTypeASt = "Numeric"
'      Case dbSingle: DeTypeASt = "Single"
      Case rdTypeVARCHAR: DeTypeASt = "Text"
      Case rdTypeDATE: DeTypeASt = "Time"
      Case rdTypeTIMESTAMP: DeTypeASt = "TimeStamp"
      Case -9: 'NVarChar
           DeTypeASt = "nvarchar"
      
      Case Else: DeTypeASt = "Desconegut"
   End Select
'       [Codi] [int] IDENTITY (1, 1) NOT NULL ,
'    [NOM] [nvarchar] (255) NULL ,
'    [PREU] [float] NOT NULL ,
'    [PreuMajor] [float] NULL ,
'    [Desconte] [float] NULL ,
'    [EsSumable] [bit] NOT NULL ,
'    [Familia] [nvarchar] (255) NULL ,
'    [CodiGenetic] [int] NOT NULL ,
'    [TipoIva] [float] NULL ,
'    [NoDescontesEspecials] [float] NULL
End Function




Sub InformaMiss(Texte As String, Optional Secundari As Boolean = False)
   DoEvents
   If Secundari Then
      frmSplash.lblVersion = Texte
   Else
      frmSplash.Estat = Texte
   End If
   DoEvents
   Debug.Print Texte
End Sub


Sub Missatges_CalEnviar(Tipus As String, Params As String, Optional CalEsborrar As Boolean = False, Optional Dbase As String = "")
   Dim Q As rdoQuery, Taula As String
   
   Taula = "MissatgesAEnviar"
   If Dbase <> "" Then Taula = Dbase & ".dbo." & Taula
   
   If Not ExisteixTaula("MissatgesAEnviar") Then ExecutaComandaSql "CREATE TABLE " & Taula & " ([Tipus] [varchar] (255) NULL ,[Param] [varchar] (255) NULL) ON [PRIMARY]"
   Set Q = Db.CreateQuery("", "Delete " & Taula & " Where Tipus = ? And Param = ? ")
   Q.rdoParameters(0) = Tipus
   Q.rdoParameters(1) = Params
   Q.Execute
   
   If Not CalEsborrar Then
      Set Q = Db.CreateQuery("", "Insert Into " & Taula & " (Tipus,Param) Values (?,?) ")
      Q.rdoParameters(0) = Tipus
      Q.rdoParameters(1) = Params
      Q.Execute
   End If
   
End Sub



Function Normalitza(s) As String
   Dim Ss As String, P As Integer
   
   Ss = ""
   If Not IsNull(s) Then Ss = s
   
   Normalitza = Ss
   P = 0
   Do
      P = InStr(P + 1, Ss, "#")
      If P > 0 Then
         Ss = Left(Ss, P) & "#" & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, Chr(13) & Chr(10))
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "#R" & Right(Ss, Len(Ss) - P - 1)
         P = P + 1
      End If
   Loop While P > 0
   
   Normalitza = Ss
   
End Function


Function NormalitzaNumero(s) As String
   Dim Ss As String, P As Integer

   Ss = ""
   If Not IsNull(s) Then Ss = s
   
   NormalitzaNumero = Ss
   
   P = 0
   Do
      P = InStr(P + 1, Ss, ",")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "." & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   NormalitzaNumero = Ss
   
End Function



Function NormalitzaDes(s As String) As String
   Dim Ss As String, P As Integer, c As String
   
   Ss = s
   NormalitzaDes = Ss
   P = 0
   Do
      P = InStr(P + 1, Ss, "##")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "#")
      If P > 0 Then
         c = ""
         If Len(Ss) >= P + 1 Then c = Mid(Ss, P + 1, 1)
         If c = "R" Then
            Ss = Left(Ss, P - 1) & Chr(13) & Chr(10) & Right(Ss, Len(Ss) - P - 1)
         End If
         P = P + 1
      End If
   Loop While P > 0
   
   NormalitzaDes = Ss
   
End Function

Sub PeticioDeConexio()
   
   ExecutaComandaSql "Use Master EXEC xp_logevent 60066,'Conexió Ara', Informational "
   
End Sub


Sub TrigguerEsborraTimeStamp(NomTaula As String)
   
   If ExisteixTaula("[Ts" & NomTaula & "]") Then ExecutaComandaSql "DROP TRIGGER [Ts" & NomTaula & "] "
 
End Sub

Public Function Obtencio_Garden_articles_sql() As Integer
    On Error GoTo Trata_error

   Dim Fil As String, f, L As String, Ll As String, sql As String, llicencia As Double
   Dim i As Integer, LlistaC As String
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, LlistaCc As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim BakL As String, K As Integer
   Dim Lll As String, FilsAImportar() As String
   Dim Pas As Integer, Valor As String, bytData() As Byte
   Dim Kk As Integer, Camps() As String, Tipus() As String
   Dim PerLine As Double, FS, Fss, FilsAImportarData() As Date
   
  ' Dim AppPath As String
   Dim Dades As String
   
   Dim Codi As String, _
    nom As String, _
    Preu As String, _
    preumajor As String, _
    desconte As String, _
    essumable As String, _
    familia As String, _
    codigenetic As String, _
    codiBarres As String, _
    tipoIva As String, _
    nodescontesespecials As String, _
    Sage_referencia As String
    
      
    CalCrearQuery = True
   ReDim Camps(0)
   Set FS = CreateObject("Scripting.FileSystemObject")
   Fil = Dir(AppPath & "\*provesarticles.SqlTrans")
   ReDim FilsAImportar(0)
   ReDim FilsAImportarData(0)
   While Len(Fil) > 0
      ReDim Preserve FilsAImportar(UBound(FilsAImportar) + 1)
      ReDim Preserve FilsAImportarData(UBound(FilsAImportarData) + 1)
      Set Fss = FS.GetFile(AppPath & "\" & Fil)
      FilsAImportar(UBound(FilsAImportar)) = Fil
      FilsAImportarData(UBound(FilsAImportarData)) = Fss.DateLastModified
      Fil = Dir
   Wend
      
   If UBound(FilsAImportar) > 0 Then     ' Copia de les taules abans de procesar
     ExecutaComandaSql "drop table [Fac_GardenPonc].[dbo].[Articles_bak]"
     ExecutaComandaSql "drop table [Fac_GardenPonc].[dbo].[Codisbarres_bak]"
     ExecutaComandaSql "select * into [Fac_GardenPonc].[dbo].[Articles_bak] from [Fac_GardenPonc].[dbo].[Articles]"
     ExecutaComandaSql "select * into [Fac_GardenPonc].[dbo].[Codisbarres_bak] from [Fac_GardenPonc].[dbo].[Codibarres]"
   End If
   
   For Kk = 1 To UBound(FilsAImportar)
      PerLine = 0
      Fil = FilsAImportar(Kk)
      f = FreeFile
    '  InformaEstat Estat, "Interpretant : Articles Sage"
      
      Dim p1, P2
      
      My_DoEvents
      
      Open AppPath & "\" & Fil For Input As #f
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" And Len(L) > 0 Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 14) = "[Sql-NomTaula:" Then
                  NomTaula = Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
            Else
      
               If CalCrearQuery Then
                  LlistaCc = ""
                  LlistaC = ""
      
                  For i = 0 To UBound(Camps)
                     LlistaC = LlistaC & "[" & Camps(i) & "]"
                     LlistaCc = LlistaCc & "?"
                     If Not i = UBound(Camps) Then
                        LlistaC = LlistaC & ","
                        LlistaCc = LlistaCc & ","
                     End If
                  Next
                  CalCrearQuery = False
              
               End If
          
              ' Inicialitzar camps de traspas
              
                codiBarres = ""
                nom = ""
                Preu = ""
                familia = ""
                Sage_referencia = ""
                tipoIva = "3"
              '----------------------------------
                 
                 
                 For i = 1 To UBound(Camps)
                  Valor = NormalitzaDes(Car(L))
On Error GoTo Trata_error
             
                   If Camps(i) = "AR_CODEBARRE" Then codiBarres = Valor
                   If Camps(i) = "AR_DESIGN" Then nom = Valor
                   If Camps(i) = "AR_PRIXVEN" Then Preu = Valor
                   If Camps(i) = "FA_CODEFAMILLE" Then familia = Valor
                   If Camps(i) = "AR_REF" Then Sage_referencia = Valor
                   If Camps(i) = "AR_PRIXACH" Then preumajor = Valor
               Next
                   
                   ' Solsament la familia del animals i llavors van al 10% d'iva
                   If familia = "ANI" Or familia = "LLAVORS4" Or familia = "LLAVORS10" Then
                   tipoIva = "2"
                   End If
                   
                   
                   Actualitza_taula_Articles nom, Preu, familia, tipoIva, codiBarres, Sage_referencia, preumajor

               PerLine = PerLine + 1
     '          If (PerLine Mod 10) = 0 Then InformaEstat Estat, "Interpretant : " & PerLine & " " & Fil
             End If
         End If
      Wend
      Close #f
      FitcherProcesat Fil
      'InformaEstat Estat, "", True
   Next
   
 Obtencio_Garden_articles_sql = 1
  ExecutaComandaSql "insert into missatgesAEnviar (tipus, param) values ('Articles', '')"
  ExecutaComandaSql "insert into missatgesAEnviar (tipus, param) values ('CodisBarres', '')"
 
Exit Function

Trata_error:
        
    Obtencio_Garden_articles_sql = 0

End Function

Function Actualitza_taula_Articles(Snom As String, _
 SPreu As String, _
 Sfamilia As String, _
 Stipoiva As String, _
 Scodi_barres As String, _
 Sreferencia As String, _
 Spreumajor As String)

Dim Codi As String, _
 nom As String, _
 Preu As String, _
 preumajor As String, _
 desconte As String, _
 essumable As String, _
 familia As String, _
 codigenetic As String, _
 tipoIva As String, _
 nodescontesespecials As String, _
 Fami1 As String, _
 Fami2 As String, _
 Fami3 As String, _
 Referencia As String, _
 Rs As rdoResultset
 Dim sql As String
    

' Posicionarse a la base de dades de Garden
' Fer una copia
' 1- Buscar article per nom
' 2 - Si el troba UPDATE
' 3 - Si no el troba INSERT

  Codi = ""
' El nom
        nom = Normalitza(Trim(Snom))
' El Preu
        If Not CStr(SPreu) = "" Then
                        Preu = NetejaNum(CStr(SPreu))
                        If Not IsNumeric(Preu) Then
                            Preu = 0
                        End If
        End If
' El Preu Major
        preumajor = "0"
        If Not CStr(Spreumajor) = "" Then
                        preumajor = NetejaNum(CStr(Spreumajor))
                        If Not IsNumeric(preumajor) Then
                            preumajor = 0
                        End If
        End If

' l,Iva
        tipoIva = Stipoiva
' La Familia

      Referencia = Normalitza(Trim(Sreferencia))
        
      If Not Sfamilia = "" Then
             Fami1 = Normalitza(Sfamilia)
             Fami2 = Normalitza(Sfamilia)
             Fami3 = Normalitza(Sreferencia)
                            
             If Fami3 = Fami2 Then Fami2 = Fami2 & "."
             If Fami3 = Fami1 Then Fami1 = Fami1 & "."
             If Fami2 = Fami1 Then Fami1 = Fami1 & "."
                            
             ExecutaComandaSql "Delete [Fac_GardenPonc].[dbo].[Families] where nom = '" & Fami3 & "' "
             ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami3 & "','" & Fami2 & "',0,3,0) "
             ExecutaComandaSql "Delete Families where nom = '" & Fami2 & "' "
             ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami2 & "','" & Fami1 & "',0,2,0) "
             ExecutaComandaSql "Delete Families where nom = '" & Fami1 & "' "
             ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami1 & "','Article',0,1,0) "
   
      End If
   
   Set Rs = Db.OpenResultset("Select codi From [Fac_GardenPonc].[dbo].[Articles] Where familia like '%" & Referencia & "%' ")
   If Rs.EOF Then
      Codi = DonamSql("Select Max(Codi) from [Fac_GardenPonc].[dbo].[Articles] ") + 1
      If Codi = "" Then Codi = 1
      sql = "Insert into [Fac_GardenPonc].[dbo].[Articles] "
      sql = sql & "(codi,nom,preu,preumajor,desconte,essumable,familia,codigenetic,tipoiva,nodescontesespecials) "
      sql = sql & "Values (" & Codi & ",'" & nom & "','" & Preu & "','" & preumajor & "',0,1,'" & Referencia & "'," & Codi & ", '" & tipoIva & "',0) "
      ExecutaComandaSql (sql)
      If Not Scodi_barres = "" Then ExecutaComandaSql "insert into [Fac_GardenPonc].[dbo].[articlespropietats]  (CodiArticle,Variable,Valor) values (" & Codi & ",'CODI_PROD','" & Scodi_barres & "') "
   Else
      Codi = Rs("codi")
      sql = "Update [Fac_GardenPonc].[dbo].[Articles]  set nom = '" & nom & "', preu = '" & Preu & "', preumajor = '" & preumajor & "',familia = '" & Referencia & "',tipoiva = '" & tipoIva & "' Where codi = " & Codi & " "
      ExecutaComandaSql (sql)
      Set Rs = Db.OpenResultset("Select valor From [Fac_GardenPonc].[dbo].[Articlespropietats] Where codiArticle = '" & Codi & "' and variable = 'CODI_PROD'")
      If Rs.EOF Then
     '   If Not Scodi_barres = "" Then ExecutaComandaSql "insert into [Fac_GardenPonc].[dbo].[articlespropietats]  (CodiArticle,Variable,Valor) values (" & Codi & ",'CODI_PROD','" & Scodi_barres & "') "
        If Not Scodi_barres = "" Then ExecutaComandaSql "insert into [Fac_GardenPonc].[dbo].[articlespropietats]  (CodiArticle,Variable,Valor) values (" & Codi & ",'CODI_PROD','" & Referencia & "') "
      Else
        'If Not Scodi_barres = "" Then ExecutaComandaSql "Update [Fac_GardenPonc].[dbo].[articlespropietats] set Valor = '" & Scodi_barres & "' where codiArticle = " & Codi & " and Variable = 'CODI_PROD' "
        If Not Scodi_barres = "" Then ExecutaComandaSql "Update [Fac_GardenPonc].[dbo].[articlespropietats] set Valor = '" & Referencia & "' where codiArticle = " & Codi & " and Variable = 'CODI_PROD' "
      End If
   End If
                 
   'Codi de barres
      If Not Scodi_barres = "" Then
        If Len(Scodi_barres) = 13 And IsNumeric(Scodi_barres) Then
           ExecutaComandaSql "Delete [Fac_GardenPonc].[dbo].[CodisBarres] Where Producte = " & Codi & " "
           ExecutaComandaSql "Delete [Fac_GardenPonc].[dbo].[CodisBarres] Where Codi = " & Scodi_barres & " "
           ExecutaComandaSql "insert into [Fac_GardenPonc].[dbo].[CodisBarres] (Producte,Codi) values (" & Codi & "," & Scodi_barres & ") "
     End If
    End If

End Function

