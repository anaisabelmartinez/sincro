Attribute VB_Name = "Fun_conversio"
Option Explicit

Global Estat
Global anterior_hora

Global Const E_lliure = 0
Global Const E_treballant = 1
Global Const E_cal_plegar = 2
Global Const E_tot_fet = 3


Const MARCA_VENTA_AL_COMPTAT = "V"
Const MARCA_CAMVI = "K"
Const MARCA_ENTRADA_DEPENDENTA = "E"
Const MARCA_PLEGA_DEP = "P"
Const MARCA_MOVIMENT_PO = "O"
Const MARCA_MOVIMENT_RA = "A"
Const MARCA_POSAR_HORA = "H"
Const MARCA_POSAR_DATA = "D"
Const MARCA_DESCONTE1 = "C"
Const MARCA_DESCONTE2 = "T"
Const MARCA_ENCARREC = "N"
Const MARCA_CAIXA_TANCADA = "Z"
Const MARCA_BON_DIA = "B"
Const MARCA_IMPRESIO_TIQUET = "I"
Const MARCA_ERROR_AL_MARCAR = "M"
Const MARCA_PA_SEC = "S"
Const MARCA_DEP_SURT_UN_MOMENT = "U"
Const MARCA_TORNA_DEPENDENTA = "Q"
Const MARCA_FACTURAT = "F"

Const OP_VENTA_AL_COMPTAT = MARCA_VENTA_AL_COMPTAT ' "Venta al comptat    "
Const OP_CAMVI = MARCA_CAMVI '"Entrat el canvi        "
Const OP_ENTRADA_DEPENDENTA = MARCA_ENTRADA_DEPENDENTA '"Arriva de dependenta"
Const OP_PLEGA_DEPENDENTA = MARCA_PLEGA_DEP '"Plega dependenta    "
Const OP_MOVIMENT_PO = MARCA_MOVIMENT_PO '"Entrat un po        "
Const OP_MOVIMENT_RA = MARCA_MOVIMENT_RA '"Entrat un ra        "
Const OP_POSAR_HORA = MARCA_POSAR_HORA '"Modificada la hora  "
Const OP_POSAR_DATA = MARCA_POSAR_DATA '"Modificada la data  "
Const OP_DESCONTE1 = MARCA_DESCONTE1 '"Fet desconte 1      "
Const OP_DESCONTE2 = MARCA_DESCONTE2 '"Fet desconte 2      "
Const OP_ENCARREC = MARCA_ENCARREC '"Fet encarrec        "
Const OP_CAIXA_TANCADA = MARCA_CAIXA_TANCADA '"Tancada la caixa    "
Const OP_BON_DIA = MARCA_BON_DIA '"Oberta la caixa     "
Const OP_IMPRESIO_TIQUET = MARCA_IMPRESIO_TIQUET '"Impres ticket       "
Const OP_ERROR_AL_MARCAR = MARCA_ERROR_AL_MARCAR '"Error al marcar     "
Const OP_PA_SEC = MARCA_PA_SEC '"Entrat pa sec       "
Const OP_DEP_SURT_UN_MOMENT = MARCA_DEP_SURT_UN_MOMENT '"Dependenta surt boti"
Const OP_TORNA_DEPENDENTA = MARCA_TORNA_DEPENDENTA '"Dependenta torna bot"
Const OP_FACTURAT = MARCA_FACTURAT '"Facturat            "
Const OP_NOMBRE_CLIENTS = "G"
Const OP_DESCUADRE = "J"
Const OP_ENTRADA_CANVI = "W"

' BASE DE DADES

Dim registre_informacio As rdoResultset

Dim botiga_en_conversio As Double
Dim data_en_conversio As Date

Dim Ultim_ticket As String

Dim u_t_botiga_en_conversio As Long
Dim u_t_data_i_hora As Date
Dim u_t_data_i_hora_UltimHuma As Date
Dim u_t_dependenta As String
Dim u_t_tick_num As Double
Dim u_t_plu As Double
Dim u_t_quantitat As Double
Dim u_t_import As Double
Dim u_t_Tipus_venta As String
Dim u_t_estat As String
Dim existeix_ultim_ticket As Boolean

Dim elTemps_NoCorre_Endarrera As Integer

Dim nom_taula_actual As String
Dim Ultima_Zeta As Date
Dim Pen_Ultima_Zeta As Date
Dim tickets_esborrats() As Double
Dim num_tickets_esborrats As Integer
Dim dos_opened As Boolean
Dim DependentesALaBotiga() As Double, NumDependentesALaBotiga As Integer
Global ParametrePerPaletaDeColors
       
Dim HoraRealZeta() As Integer

Global DIRECTORI_TEMPORAL As String
Global DIRECTORI_TEMPORAL_BAK As String
Global DIRECTORI_TEMPORAL_ERR As String

Dim PreuProducte As rdoQuery

Dim Registre_Tickets    As rdoQuery
Dim Registre_Movimen    As rdoQuery
Dim Registre_Horaris    As rdoQuery
Dim Registre_Pa_Sec     As rdoQuery
Dim Esb1                As rdoQuery
Dim Esb2                As rdoQuery

Dim HashPreus(9999) As Double
Dim ControlEstat  As Label
Dim FilActual As String
Dim Q_ExisteixTaula As rdoQuery
Dim Handel1, Handel2
Sub AjuntaBotiguesSiCal()
'   If Not ExisteixTaula("ParamsHw") Then
'      Sql = "CREATE TABLE [FeinesAFer] ("
'      Sql = Sql & " [Id]      [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Tipus]   [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Ciclica] [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Param1]  [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Param2]  [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Param3]  [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Param4]  [nvarchar] (255)   NULL ,"
'      Sql = Sql & " [Param5]  [nvarchar] (255)   NULL"
'      Sql = Sql & " ) ON [PRIMARY]"
'      ExecutaComandaSql Sql
'      Sql = " ALTER TABLE [dbo].[FeinesAFer] WITH NOCHECK ADD "
'      Sql = Sql & " CONSTRAINT [DF_FeinesAFer_Id] DEFAULT (newid()) FOR [Id] "
'      ExecutaComandaSql Sql
'   End If

    
    
End Sub


Function DonamNomTaulaAsientosContables(dia As Date) As String
    Dim sql
    
   DonamNomTaulaAsientosContables = "AsientosContables_" & Year(dia)
   
   If Not ExisteixTaula(DonamNomTaulaAsientosContables) Then
        sql = "CREATE TABLE " & DonamNomTaulaAsientosContables & " ("
        sql = sql & "[idNorma43]          [nvarchar](255) NOT NULL, "
        sql = sql & "[orden]              [int] NOT NULL, "
        sql = sql & "[FECHA]              [datetime] NULL, "
        sql = sql & "[C]                  [nvarchar](255) NULL, "
        sql = sql & "[CONCEPTO]           [nvarchar](255) NULL, "
        sql = sql & "[DESCRIPCION]        [nvarchar](255) NULL, "
        sql = sql & "[N_DOC]              [nvarchar](255) NULL, "
        sql = sql & "[SUBCUENTA]          [nvarchar](255) NULL, "
        sql = sql & "[DEBE]               [float] NULL, "
        sql = sql & "[HABER]              [float] NULL, "
        sql = sql & "[FacturaId]          [nvarchar](255) NULL, "
        sql = sql & "[FacturaNum]         [nvarchar](255) NULL, "
        sql = sql & "[TaulaFactura]       [nvarchar](255) NULL, "
        sql = sql & "[IVA]                [nvarchar](255) NULL, "
        sql = sql & "[T_IVA]              [nvarchar](255) NULL, "
        sql = sql & "[NumRegistre]        [nvarchar](255) NULL, "
        sql = sql & "[Revisat]            [nvarchar](255) NULL, "
        sql = sql & "[ReferenciaInterna]  [nvarchar](255) NULL, "
        sql = sql & "[DescripcionInterna] [nvarchar](255) NULL, "
        sql = sql & "[FECHA_OPERACION]    [datetime] NULL "
        sql = sql & ") ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
   End If

End Function


Function DonamNomTaulaAlbaransNum(dia As Date) As String
   Dim sql
    
   DonamNomTaulaAlbaransNum = "[AlbaransNum_" & Year(dia) & "]"
   
   If Not ExisteixTaula(DonamNomTaulaAlbaransNum) Then
        sql = "CREATE TABLE " & DonamNomTaulaAlbaransNum & " ("
        sql = sql & " [id]       [nvarchar] (255)   NULL Default (NEWID()),"
        sql = sql & " [tipus]    [nvarchar] (255)   NULL,"
        sql = sql & " [Codi]     [numeric]  (18,0)  IDENTITY (1,1) NOT FOR REPLICATION  NOT NULL,"
        sql = sql & " [Data]     [datetime]         NULL,"
        sql = sql & " [Viatge]   [nvarchar] (255)   NULL,"
        sql = sql & " [Client]   [numeric]  (18,0)  NULL,"
        sql = sql & " [Import]   [float]            NULL,"
        sql = sql & " [Validado] [nvarchar] (255)   NULL,"
        sql = sql & " [Libre]    [nvarchar] (255)   NULL"
        sql = sql & " ) ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
        ExecutaComandaSql "Insert Into " & DonamNomTaulaAlbaransNum & " (Data,Viatge,Client) values(getdate(), 'Inicial',0)"
   End If

End Function


Function DonamNomTaulaHistoricServit(dia As Date) As String
    Dim sql
    
   DonamNomTaulaHistoricServit = "HistoricServit_" & Year(dia) & "_" & Right("00" & Month(dia), 2)
   
   If Not ExisteixTaula(DonamNomTaulaHistoricServit) Then
        sql = "CREATE TABLE " & DonamNomTaulaHistoricServit & " ("
        sql = sql & "[Modificat]         [datetime] Default (GetDate()), "
        sql = sql & "[IdServit]          [nvarchar] (255) Null, "
        sql = sql & "[DataServit]        [datetime]       Null, "
        sql = sql & "[Tipus]             [nvarchar] (5)   Null, "   'ALTA, MODIFICACIÓ, DELETE
        sql = sql & "[Desde]             [nvarchar] (25)  Null, "   'GDT, RECORDA, TPV, TPVINFORMA, SINCRO, ...
        sql = sql & "[QuiStamp]          [nvarchar] (255) Null, "
        sql = sql & "[Client]            [float]          Null, "
        sql = sql & "[CodiArticle]       [int]            Null, "
        sql = sql & "[Viatge]            [nvarchar] (255) Null, "
        sql = sql & "[Equip]             [nvarchar] (255) Null, "
        sql = sql & "[QuantitatDemanada] [float]    Default (0), "
        sql = sql & "[QuantitatTornada]  [float]    Default (0), "
        sql = sql & "[QuantitatServida]  [float]    Default (0), "
        sql = sql & "[TipusComanda]      [float]          Null, "
        sql = sql & "[Comentari]         [nvarchar] (255) Null "
        sql = sql & ") ON [PRIMARY]"
        
        ExecutaComandaSql (sql)
   End If
   
End Function

Function DonamTaulaServit(dia As Date) As String
    Dim NomTaula As String
    
    NomTaula = "Servit-" & Format(dia, "yy-mm-dd")
    DonamTaulaServit = LastDatabase & ".dbo.[" & NomTaula & "]"

    If Not ExisteixTaula(DonamTaulaServit) Then CreaTaulaServit2 NomTaula
End Function

Function DonamTaulaServit2(dia As Date, Database As String) As String
    Dim NomTaula As String
    
    NomTaula = "Servit-" & Format(dia, "yy-mm-dd")
    DonamTaulaServit2 = Database & ".dbo.[" & NomTaula & "]"

    If Not ExisteixTaula(DonamTaulaServit2) Then CreaTaulaServit2 NomTaula
End Function

Sub InsertFeineaAFer(Tipus As String, Optional p1 As String, Optional P2 As String, Optional P3 As String, Optional P4 As String, Optional P5 As String)
   Dim Q As rdoQuery, sql As String
   
   If Not ExisteixTaula("FeinesAFer") Then
      sql = "CREATE TABLE [FeinesAFer] ("
      sql = sql & " [Id]      [nvarchar] (255)   NULL ,"
      sql = sql & " [Tipus]   [nvarchar] (255)   NULL ,"
      sql = sql & " [Ciclica] [nvarchar] (255)   NULL ,"
      sql = sql & " [Param1]  [nvarchar] (255)   NULL ,"
      sql = sql & " [Param2]  [nvarchar] (255)   NULL ,"
      sql = sql & " [Param3]  [nvarchar] (255)   NULL ,"
      sql = sql & " [Param4]  [nvarchar] (255)   NULL ,"
      sql = sql & " [Param5]  [nvarchar] (255)   NULL, "
      sql = sql & " [tmStmp]  [datetime]         NULL CONSTRAINT [DF_FeinesAFer_TmStmp] DEFAULT (getdate()) "
      sql = sql & " ) ON [PRIMARY]"
      ExecutaComandaSql sql
      sql = " ALTER TABLE [dbo].[FeinesAFer] WITH NOCHECK ADD "
      sql = sql & " CONSTRAINT [DF_FeinesAFer_Id] DEFAULT (newid()) FOR [Id] "
      ExecutaComandaSql sql
   End If
   
   Set Q = Db.CreateQuery("", "Insert Into FeinesAFer ([Tipus] , [Ciclica], [Param1], [Param2], [Param3], [Param4], [Param5]) Values (?,?,?,?,?,?,?)")
   Q.rdoParameters(0) = Tipus
   Q.rdoParameters(1) = 0
   Q.rdoParameters(2) = p1
   Q.rdoParameters(3) = P2
   Q.rdoParameters(4) = P3
   Q.rdoParameters(5) = P4
   Q.rdoParameters(6) = P5
   Q.Execute
   
End Sub

Sub DestriaCamps(Ll As String, Camps() As String, Tipus() As String)
   
   ReDim Camps(0)
   ReDim Tipus(0)
   
   While Len(Ll) > 0
      ReDim Preserve Camps(UBound(Camps) + 1)
      ReDim Preserve Tipus(UBound(Tipus) + 1)
      Camps(UBound(Camps)) = Car(Ll)
      Tipus(UBound(Tipus)) = Car(Ll)
      If UCase(Tipus(UBound(Tipus))) = "DOUBLE" Then Tipus(UBound(Tipus)) = "Numeric"
   Wend

End Sub

Sub InsertFeinaAfer(p1 As String, P2, P3, P4, P5)

   

End Sub

Sub InterpretaFitchersVentes(Optional Estat As Label = Nothing)
      
   Set ControlEstat = Estat
   CarregaHashPreus
   My_DoEvents
   
   Interpreta_SqlTrans_Moviments Estat
   Interpreta_SqlTrans_Salutacions Estat
   Interpreta_SqlTrans_Venut Estat
   Interpreta_SqlTrans_Albarans Estat
   Interpreta_SqlTrans_Anulat Estat
   Interpreta_SqlTrans_Tornat Estat
   Interpreta_SqlTrans_Inventari Estat
   Interpreta_SqlTrans_Log Estat
   Interpreta_SqlTrans_Encarregs Estat
   Interpreta_SqlTrans_Caixa Estat
      
End Sub

Sub Interpreta_SqlTrans_Caixa(Optional Estat As Label = Nothing)
   Dim Fil As String, f, L As String, Ll As String, Contingut As String, NomMaquina As String, nomfile As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double
   Dim Tipus_Moviment As String, import As Double, Motiu As String
   Dim Cm_Primera As String, Cm_Oberta As String, Caca As String, Cm_Tancada As String
   Dim Cm_CanviInicial As String, Cm_CanviInicialDetall  As String, Cm_CanviFinal As String, Cm_CanviFinalDetall As String, Cm_PrimerTicket As String, Cm_UltimTicket As String
   Dim Cm_HInici As Date, Cm_HFi As Date, Cm_Zeta As Double, Cm_descuadre As Double, Cm_Clients As Double
   Dim Ca As String, P As Integer, VCaca As Double
   Dim sD As String, Camps() As String, Tipus() As String
   Dim rsOracul As rdoResultset
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*Caixa.SqlTrans")
   
   On Error GoTo norR
'   On Error GoTo 0
   
   While Len(Fil) > 0
   
    If (ExisteixTaula("SFPT_DEBUG")) Then
    ExecutaComandaSql "insert into SFPT_DEBUG values (getdate(), 'INTERPRETA CAIXA', '', '', '', '', '" & Right(Fil, 250) & "', '')"
    End If
   
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Caixa : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
               If Left(L, 17) = "[Sql-LlistaCamps:" Then
                  CalCrearQuery = True
                  DestriaCamps Ll, Camps, Tipus
               End If
               
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaMovi(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? And (Tipus_Moviment='W' Or Tipus_Moviment='Wi' Or Tipus_Moviment='G' Or Tipus_Moviment='J' Or Tipus_Moviment='Z' ) ")
'                  Q1.rdoParameters(0) = Botiga
'                  Q1.rdoParameters(1) = DataInici
'                  Q1.rdoParameters(2) = DataFi
'                  Q1.Execute
                  
                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaMovi(dataInici) & "] (Botiga,Data,DependentA,Tipus_Moviment,Import,Motiu) Values(?,?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
               End If
On Error GoTo norR
               Cm_Zeta = 0
               Cm_descuadre = 0
               Cm_Clients = 0
               Ll = L
               For i = 1 To UBound(Camps)
                  Caca = Car(L)
                  VCaca = 0
                  If IsNumeric(Caca) Then VCaca = Val(Caca)
                  Select Case Camps(i)
                     Case "Primera":            Cm_Primera = Caca
                     Case "Oberta":             Cm_Oberta = Caca
                     Case "Tancada":            Cm_Tancada = Caca
                     Case "CanviInicial":       Cm_CanviInicial = VCaca
                     Case "CanviInicialDetall": Cm_CanviInicialDetall = Caca
                     Case "CanviFinal":         Cm_CanviFinal = VCaca
                     Case "CanviFinalDetall":   Cm_CanviFinalDetall = Caca
                     Case "PrimerTicket":       Cm_PrimerTicket = VCaca
                     Case "UltimTicket":        Cm_UltimTicket = VCaca
                     Case "HInici":             If IsDate(Caca) Then Cm_HInici = CDate(Caca) Else Cm_HInici = "0:00:00"
                     Case "HFi":                If IsDate(Caca) Then Cm_HFi = CDate(Caca) Else Cm_HFi = "0:00:00"
                     Case "Zeta":               Cm_Zeta = VCaca
                     Case "Descuadre":          Cm_descuadre = VCaca
                     Case "Clients":            Cm_Clients = VCaca
                  End Select
               Next
               
               If Cm_Tancada Then
                  Dim incremental As Double
                  
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = Cm_HInici
                  Q1.rdoParameters(2) = Cm_HFi
                  Q1.Execute
                  
                  incremental = 0
                  Do
                     If Cm_CanviInicialDetall = "" Or Cm_CanviInicialDetall = "0" Then Cm_CanviInicialDetall = "[0.01;0][0.02;0][0.05;0][0.1;0][0.2;0][0.5;0][1;0][2;0][5;0][10;0][20;0][50;0][100;0][500;0]][90.58][[0.01;0.29][0.02;4.14][0.05;16.85][0.1;16.9][0.2;0.4][0.5;2][1;48][2;2][5;0][10;0][20;0][50;0][100;0][500;0]"
                     Ca = Car(Cm_CanviInicialDetall)
                     If Len(Ca) > 0 Then
                        P = InStr(Ca, ";")
                        If P > 1 Then
                           Motiu = Left(Ca, P - 1)
                           import = Right(Ca, Len(Ca) - P)
                           If incremental > Val(Motiu) Then Exit Do
                           incremental = Val(Motiu)
                           Q2.rdoParameters(1) = Cm_HInici
                           Q2.rdoParameters(2) = Cm_Primera
                           Q2.rdoParameters(3) = "Wi"
                           Q2.rdoParameters(4) = import
                           Q2.rdoParameters(5) = "En : " & Motiu
                           Q2.Execute
                        End If
                     End If
                  Loop While Len(Ca) > 0
               
                  incremental = 0
                  Do
                     If Cm_CanviFinalDetall = "" Or Cm_CanviFinalDetall = "0" Then Cm_CanviFinalDetall = "[0.01;0][0.02;0][0.05;0][0.1;0][0.2;0][0.5;0][1;0][2;0][5;0][10;0][20;0][50;0][100;0][500;0]][90.58][[0.01;0.29][0.02;4.14][0.05;16.85][0.1;16.9][0.2;0.4][0.5;2][1;48][2;2][5;0][10;0][20;0][50;0][100;0][500;0]"
                     Ca = Car(Cm_CanviFinalDetall)
                     If Len(Ca) > 0 Then
                        P = InStr(Ca, ";")
                        If P > 1 Then
                           Motiu = Left(Ca, P - 1)
                           import = Right(Ca, Len(Ca) - P)
                           If incremental > Val(Motiu) Then Exit Do
                           incremental = Val(Motiu)
                           Q2.rdoParameters(1) = Cm_HFi
                           Q2.rdoParameters(2) = Cm_Primera
                           Q2.rdoParameters(3) = "W"
                           Q2.rdoParameters(4) = import
                           Q2.rdoParameters(5) = "En : " & Motiu
                           Q2.Execute
                        End If
                     End If
                  Loop While Len(Ca) > 0
               
                  Q2.rdoParameters(1) = Cm_HFi  '"J"  Descuadre
                  Q2.rdoParameters(2) = Cm_Primera
                  Q2.rdoParameters(3) = "J"
                  Q2.rdoParameters(4) = Cm_descuadre
                  Q2.rdoParameters(5) = ""
                  Q2.Execute
                  Q2.rdoParameters(1) = Cm_HFi  '"G"  Clients
                  Q2.rdoParameters(2) = Cm_Primera
                  Q2.rdoParameters(3) = "G"
                  Q2.rdoParameters(4) = Cm_Clients
                  Q2.rdoParameters(5) = ""
                  Q2.Execute
                  Q2.rdoParameters(1) = Cm_HFi  '"Z"  Zeta
                  Q2.rdoParameters(2) = Cm_Primera
                  Q2.rdoParameters(3) = "Z"
                  Q2.rdoParameters(4) = Cm_Zeta
                  Q2.rdoParameters(5) = ""
                  Q2.Execute

                  InsertFeineaAFer "VigilarAlertes", "Caixa", "[" & Format(Cm_HInici, "dd-mm-yyyy ") & " de " & Format(Cm_HInici, "hh:mm") & " a " & Format(Cm_HFi, "hh:mm") & "]", "[" & botiga & "]", CStr(Cm_descuadre), CStr(Cm_Zeta)
                                    
                  'Sincronitzacio MURANO
                  If Cm_Zeta > 0 Then
                    If getEmpresaMURANOBotiga(botiga) <> -1 Then
                        InsertFeineaAFer "SincroMURANOCaixaOnLine", "[" & botiga & "]", "[" & Format(Cm_HInici, "dd-mm-yyyy") & " " & Format(Cm_HInici, "hh:mm:ss") & "]", "[" & Format(Cm_HFi, "dd-mm-yyyy") & " " & Format(Cm_HFi, "hh:mm:ss") & "]", "[" & Cm_PrimerTicket & "," & Cm_UltimTicket & "]", "[" & Cm_Zeta & "]"
                    End If
                  End If
                  
                  'Cáculo previsiones
                  If (ExisteixTaula("ORACUL_L_PARAMS")) And Cm_Zeta > 0 Then
                        Dim serie As String
                    
                        serie = "MATI"
                        If DatePart("h", Cm_HFi) > 15 Then serie = "TARDA"
                    
                        Set rsOracul = Db.OpenResultset("SELECT * FROM [ORACUL_L_PARAMS] WHERE BOTIGA = " & botiga & " and id = 'BASE_" & serie & "'")
                        If Not rsOracul.EOF Then
                            ExecutaComandaSql "insert into feinesafer values (newid(), 'PrevisionsVendesDiari', 0, '" & Format(Cm_HInici, "dd/mm/yyyy") & "', '" & botiga & "', '" & serie & "', '' ,'', GETDATE())"
                        End If
                  End If
               End If
            End If
         End If
      Wend

    If (ExisteixTaula("SFPT_DEBUG")) Then
    ExecutaComandaSql "insert into SFPT_DEBUG values (getdate(), 'SIN_ERROR', '" & botiga & "', '', '', '', '" & Right(Fil, 250) & "', '" & Left(err.Description, 250) & "')"
    End If

      TancaFile f
      FitcherProcesat Fil
Mas:
      Fil = Dir
   Wend
   
   Exit Sub
   
norR:
    If (ExisteixTaula("SFPT_DEBUG")) Then
    ExecutaComandaSql "insert into SFPT_DEBUG values (getdate(), 'ERROR_INTERPRETA', '" & botiga & "', '', '', '', '" & Right(Fil, 250) & "', '" & Left(err.Description, 250) & "')"
    End If
    
   TancaFile f
   FitcherProcesat Fil, False
   GoTo Mas
End Sub


Sub Interpreta_SqlTrans_Venut(Optional Estat As Label = Nothing)
    Dim Contingut As String, NomMaquina As String, nomfile      As String
    Dim Fil As String, f, L As String, Ll As String
    Dim Codi As Double, Preu As Double, nom As String
    Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
    Dim LlistaCamps As String, NomTaula As String, Lll As String
    Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
    Dim data As Date, dependenta As Double, Caca As String
    Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
    Dim rs As rdoResultset
    CalCrearQuery = True
    
    Fil = Dir(AppPath & "\*Venut.SqlTrans")
    'On Error GoTo Noor
   
    While Len(Fil) > 0
        f = FreeFile
        If Not Estat Is Nothing Then Estat.Caption = "Interpretant Ventas : " & Fil
        My_DoEvents
        Open AppPath & "\" & Fil For Input As #f
        DescomposaContingut Fil, Contingut, NomMaquina, nomfile
        While Not EOF(f)
           Line Input #f, L
           If Len(L) > 0 And Not Left(L, 1) = "#" Then
              If Left(L, 5) = "[Sql-" Then
                 Ll = DonamParam(L)
                 If Left(L, 15) = "[Sql-DataInici:" Then
                    CalCrearQuery = True
                    dataInici = CVDate(Ll)
                 End If
                 If Left(L, 12) = "[Sql-DataFi:" Then
                    CalCrearQuery = True
                    dataFi = CVDate(Ll)
                 End If
                 If Left(L, 12) = "[Sql-Botiga:" Then
                    CalCrearQuery = True
                    botiga = Ll
                    If Not NomMaquina = "" Then LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
                 End If
                 If Left(L, 14) = "[Sql-NomTaula:" Then
                    If InStr(1, L, "Venut") = 0 Then
                        GoTo SaltaFile
                    End If
                 End If
              Else
                 If CalCrearQuery Then
                    CreaTaulesDadesTpv dataInici
                    Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaVentas(dataInici) & "] Where Botiga = ? and (Estat = ? or Estat = '' or Estat is null) And Data >= ? And Data <= ? ")
                    Q1.rdoParameters(0) = botiga
                    Q1.rdoParameters(1) = botiga
                    Q1.rdoParameters(2) = dataInici
                    Q1.rdoParameters(3) = dataFi
                    Q1.Execute
        'If UCase(EmpresaActual) = UCase("Tena") Then ExecutaComandaSql "insert into debug values (getdate(), 'Delete [" & NomTaulaVentas(dataInici) & "] Where Botiga = " & botiga & " and (Estat = " & botiga & " or Estat = '' or Estat is null) And Data >= " & dataInici & " And Data <= " & dataFi & "')"
        
                    Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaVentas(dataInici) & "] Where Estat = ? And Data >= ? And Data <= ? ")
                    Q1.rdoParameters(0) = botiga
                    Q1.rdoParameters(1) = dataInici
                    Q1.rdoParameters(2) = dataFi
                    Q1.Execute
        'If UCase(EmpresaActual) = UCase("Tena") Then ExecutaComandaSql "insert into debug values (getdate(), 'Delete [" & NomTaulaVentas(dataInici) & "] Where Estat = " & botiga & " And Data >= " & dataInici & " And Data <= " & dataFi & "')"
        
                    Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaVentas(dataInici) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar) Values (?,?,?,?,?,?,?,?,?,?,?) ")
                    Q2.rdoParameters(0) = botiga
                    CalCrearQuery = False
                 End If
                 Lll = L
                 On Error Resume Next
                 Cm_HInici = Car(L)
                 Cm_Dependenta = Car(L)
                 Cm_NumTic = Car(L)
                 Caca = Car(L)  ' Botiga
                 Cm_Article = Car(L)
        '               Cm_Article = ArticleCodiGeneticCodi(Cm_Article)
                 Cm_Quantitat = Val(Car(L))
                 Cm_Preu = Val(Car(L))
                 Cm_Import = Val(Car(L))
                 Cm_Descompte = Car(L)
                 Cm_Origen = Car(L)
                 Cm_Otros = Car(L)
                 
                 Q2.rdoParameters(1) = Cm_HInici
                 Q2.rdoParameters(2) = Cm_Dependenta
                 Q2.rdoParameters(3) = Cm_NumTic
                 Q2.rdoParameters(4) = ""
                 Q2.rdoParameters(5) = Cm_Article
                 Q2.rdoParameters(6) = Cm_Quantitat
                 Q2.rdoParameters(7) = Cm_Import
                 If Cm_Descompte = 0 Then
                    Q2.rdoParameters(8) = "V"
                 Else
                    Q2.rdoParameters(8) = "Desc_" & Cm_Descompte
                 End If
                 Q2.rdoParameters(9) = Cm_Otros
                 Q2.rdoParameters(10) = Cm_Origen
                 Q2.Execute
                 On Error GoTo Noor
              End If
           End If
        Wend
        
        ' AjuntaBotiguesSiCal DataInici
        InsertFeineaAFer "AjuntaTpvS", "[" & Year(dataInici) & "]", "[" & Month(dataInici) & "]", "[" & Day(dataInici) & "]", "[" & botiga & "]"
        InsertFeineaAFer "AcumulaPunts", "Ventes", "[" & Year(dataInici) & "]", "[" & Month(dataInici) & "]"
        InsertFeineaAFer "CalcularPrevisions", "Ventes", "[" & Format(dataInici, "dd-mm-yyyy") & " de " & Format(dataInici, "hh:mm") & " a " & Format(dataFi, "hh:mm") & "]", "[" & botiga & "]"
        InsertFeineaAFer "VigilarAlertes", "Ventes", "[" & Format(dataInici, "dd-mm-yyyy") & " de " & Format(dataInici, "hh:mm") & " a " & Format(dataFi, "hh:mm") & "]", "[" & botiga & "]"
        InsertFeineaAFer "FesComandes", "PreDefinides", "[" & Format(DateAdd("d", 6, dataInici), "dd-mm-yyyy") & "]", "[" & botiga & "]", "[" & Format(DateAdd("d", 1, dataInici), "dd-mm-yyyy") & "]"
        InsertFeineaAFer "CalculaEstadistic", "Ventes", "[" & Format(dataInici, "dd-mm-yyyy") & "]", "[" & botiga & "]"
        InsertFeineaAFer "CalcularObjectius", "[" & Format(DateAdd("d", 7, dataInici), "dd-mm-yyyy") & "]", "[" & Format(DateAdd("d", 7, dataFi), "dd-mm-yyyy") & "]", "[" & botiga & "]"
        If UCase(EmpresaActual) = UCase("Tena") Then
            'Le pongo una A_ para que se ejecute después del CalcularPrevisions (va por orden inverso de tipus)
            InsertFeineaAFer "A_RecalcularTiquetsPromocions", "[" & Format(dataInici, "dd/mm/yyyy") & " " & Format(dataInici, "hh:mm:ss") & "]", "[" & Format(dataFi, "dd/mm/yyyy") & " " & Format(dataFi, "hh:mm:ss") & "]", "[" & botiga & "]"
        End If
                
SaltaFile:
                
        TancaFile f
        FitcherProcesat Fil
      'ExecutaComandaSql "insert into debug values (getdate(), 'FITXER VENDES PROCESSAT: " & nomfile & " ERR: " & err.Description & "')"
Mas:
        Fil = Dir
    Wend
   
   Exit Sub
   
Noor:

   TancaFile f
   FitcherProcesat Fil, False
'ExecutaComandaSql "insert into debug values (getdate(), 'ERROR FITXER VENDES: " & nomfile & " ERR: " & err.Description & "')"

   Resume Mas
End Sub

Sub Interpreta_SqlTrans_RecepcioMP(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile As String
   Dim Fil As String, f, L As String, Ll As String, sql As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String, Lll As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double
   Dim Cm_Dependenta As Double, Cm_EAN As Double, Cm_Botiga As Double, Cm_Id As Double, Cm_HInici As Date
   Dim rs As rdoResultset
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*RecepcioMateriesPrimeres.SqlTrans")
   'On Error GoTo Noor
   
   While Len(Fil) > 0
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Recepcio Materies Primeres : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  If Not NomMaquina = "" Then LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  If Not ExisteixTaula("[" & NomTaulaRecepcioMP(dataInici) & "]") Then
                     sql = "[TmSt][DateTime], [Dependenta][float],  [Botiga][nvarchar] (255), [EAN][nvarchar] (255), [Id][nvarchar] (255)"
                     ExecutaComandaSql "Create Table [" & NomTaulaRecepcioMP(dataInici) & "] (" & sql & ") "
                  End If

                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaRecepcioMP(dataInici) & "] (Botiga,TmSt,Dependenta,EAN,Id) Values (?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
               End If
               Lll = L
               On Error Resume Next
               Cm_HInici = Car(L)
               Cm_Dependenta = Car(L)
               Cm_Botiga = Car(L)  ' Botiga
               Cm_EAN = Car(L)
               Cm_Id = Car(L)
               
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = Cm_Dependenta
               Q2.rdoParameters(3) = Cm_EAN
               Q2.rdoParameters(4) = Cm_Id
               Q2.Execute
               On Error GoTo Noor
            End If
         End If
      Wend
      
      TancaFile f
      FitcherProcesat Fil
      
Mas:
      Fil = Dir
   Wend
   Exit Sub
   
Noor:

   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub
Sub Interpreta_SqlTrans_Albarans(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer, SqlExporta1 As String, SqlExporta2 As String
   Dim data As Date, dependenta As Double, Caca As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*Albarans.SqlTrans")
   On Error GoTo Noor
   
   While Len(Fil) > 0
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Albarans : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      SqlExporta1 = ""
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaAlbarans(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute

                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaAlbarans(dataInici) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar) Values (?,?,?,?,?,?,?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  SqlExporta1 = "[" & Format(dataInici, "dd-mm-yyyy") & " de " & Format(dataInici, "hh:mm") & " a " & Format(dataFi, "hh:mm") & "]"
                  SqlExporta2 = "[" & botiga & "]"
                  CalCrearQuery = False
               End If
               Cm_HInici = Car(L)
               Cm_Dependenta = Car(L)
               Cm_NumTic = Car(L)
               Caca = Car(L)  ' Botiga
               Cm_Article = Car(L)
'               Cm_Article = ArticleCodiGeneticCodi(Cm_Article)
               Cm_Quantitat = Val(Car(L))
               Cm_Preu = Val(Car(L))
               Cm_Import = Val(Car(L))
               Cm_Descompte = Car(L)
               Cm_Origen = Car(L)
               Cm_Otros = Car(L)
               
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = Cm_Dependenta
               Q2.rdoParameters(3) = Cm_NumTic
               Q2.rdoParameters(4) = ""
               Q2.rdoParameters(5) = Cm_Article
               Q2.rdoParameters(6) = Cm_Quantitat
               Q2.rdoParameters(7) = Cm_Import
               If Cm_Descompte = 0 Then
                  Q2.rdoParameters(8) = "V"
               Else
                  Q2.rdoParameters(8) = "Desc_" & Cm_Descompte
               End If
               Q2.rdoParameters(9) = Cm_Otros
               Q2.rdoParameters(10) = Cm_Origen
               Q2.Execute
            End If
         End If
      Wend
      If Len(SqlExporta1) > 0 Then
         ExportaAlbarans SqlExporta1, SqlExporta2
         InsertFeineaAFer "ExportaAlbarans", "DeBotiga", SqlExporta1, SqlExporta2
         SqlExporta1 = ""
      End If
      TancaFile f
      FitcherProcesat Fil
      
Mas:
      Fil = Dir
   Wend
   Exit Sub
   
Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub

Sub Interpreta_SqlTrans_Anulat(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As Double, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*Anulats.SqlTrans")
   On Error GoTo Noor
   
   While Len(Fil) > 0
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Anulats : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaAnulats(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute

                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaAnulats(dataInici) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,Otros,FormaMarcar) Values(?,?,?,?,?,?,?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
' Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta
               End If
Ll = L
               Cm_HInici = Car(L)
               Cm_Dependenta = Car(L)
               Cm_NumTic = Car(L)
               Caca = Car(L)  ' Botiga
               Cm_Article = Car(L)
               Cm_Quantitat = Car(L)
               Cm_Preu = Car(L)
               Cm_Import = Car(L)
               Cm_Descompte = Car(L)
               Cm_Origen = Car(L)
               Cm_Otros = Car(L)
               
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = Cm_Dependenta
               Q2.rdoParameters(3) = Cm_NumTic
               Q2.rdoParameters(4) = ""
               Q2.rdoParameters(5) = Cm_Article
               Q2.rdoParameters(6) = Cm_Quantitat
               Q2.rdoParameters(7) = Cm_Import
               If Cm_Descompte = 0 Then
                  Q2.rdoParameters(8) = "V"
               Else
                  Q2.rdoParameters(8) = "Desc_" & Cm_Descompte
               End If
               Q2.rdoParameters(9) = Cm_Otros
               Q2.rdoParameters(10) = Cm_Origen
               Q2.Execute
            End If
         End If
      Wend
      TancaFile f
      FitcherProcesat Fil
Mas:
      Fil = Dir
   Wend
   Exit Sub
   
Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub

Sub Interpreta_SqlTrans_Tornat(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String, NoDevolucions As Boolean, Rs5 As rdoResultset, Fil As String, f, L As String, Ll As String, Boti() As Double
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery, Codi As Double, Preu As Double, nom As String, D As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer, LlistaCamps As String, NomTaula As String, LlistaFiles() As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As String, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date, data As Date, dependenta As Double, Caca As String, Sa As String, aa, bb, cc
   
   ReDim LlistaFiles(0)
   D = Dir(AppPath & "\*Tornat.SqlTrans")
   While Len(D) > 0
      ReDim Preserve LlistaFiles(UBound(LlistaFiles) + 1)
      LlistaFiles(UBound(LlistaFiles)) = D
      D = Dir
   Wend

On Error GoTo 0
On Error GoTo Noor
   
   For i = 1 To UBound(LlistaFiles)
      Fil = LlistaFiles(i)
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Pa Sec : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      CalCrearQuery = True
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                    CreaTaulesDadesTpv dataInici
                    Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaDevol(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? ")
                    Q1.rdoParameters(0) = botiga
                    Q1.rdoParameters(1) = dataInici
                    Q1.rdoParameters(2) = dataFi
                    Q1.Execute
                    Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaDevol(dataInici) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta) Values(?,?,?,?,?,?,?,?,?) ")
                    Q2.rdoParameters(0) = botiga
                    Q2.rdoParameters(3) = 0
                    Q2.rdoParameters(4) = ""
                    Q2.rdoParameters(8) = "S"
                    CalCrearQuery = False
               End If
               Ll = L
               Cm_HInici = Car(L)
               Caca = Car(L)  ' Botiga
               Cm_Article = Car(L)
               Cm_Quantitat = Car(L)
               Caca = Car(L)  ' TipusDevolucio
               Q2.rdoParameters(5) = Cm_Article
               Q2.rdoParameters(6) = Cm_Quantitat
               Q2.rdoParameters(7) = ArticleCodiPvp(ArticleCodiGeneticCodi(Cm_Article)) * Cm_Quantitat
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = 0   'Dependenta

               
               Q2.Execute
            End If
         End If
      
      Wend
      On Error Resume Next
         Q2.Close
         Q1.Close
         CalCrearQuery = True
      On Error GoTo Noor
      TancaFile f
      FitcherProcesat Fil
      'If Not TePermisPer("NoAvonarDevolucions") Then
      
      NoDevolucions = False
      Set Rs5 = Db.OpenResultset("select Valor from constantsclient where codi = " & botiga & " and variable= 'NoDevolucions' ")
      If Not Rs5.EOF Then If Not IsNull(Rs5(0)) Then If IsNumeric(Rs5(0)) Then NoDevolucions = Len(Rs5(0)) > 0
      
      If Not NoDevolucions Then ActualizaPanSeco dataInici, CDbl(botiga)
      
Mas:
   Next
   
   
'For i = 0 To 15
'   DataInici = DateAdd("d", -i, Now)
'   CarregaLlistaBotigues "00", Boti, Year(DataInici), Month(DataInici), 1, Day(DataInici)
'   Dim j
'   For j = 0 To UBound(Boti)
'       Botiga = Boti(j)
'       NoDevolucions = False
'       Set Rs5 = Db.OpenResultset("select Valor from constantsclient where codi = " & Botiga & " and variable= 'NoDevolucions' ")
'       If Not Rs5.EOF Then If Not IsNull(Rs5(0)) Then If IsNumeric(Rs5(0)) Then NoDevolucions = Len(Rs5(0)) > 0
'       If Not NoDevolucions Then ActualizaPanSeco DataInici, CDbl(Botiga)
'   Next
'Next
   
   Exit Sub

Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub

Sub Interpreta_SqlTrans_Inventari(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String, SqlExporta1 As String, SqlExporta2 As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String, D As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String, LlistaFiles() As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, Sa As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As String, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim TipusVenta As String, Proveidor As String
   Dim aa, bb, cc, DataTot As String
   
   
   ReDim LlistaFiles(0)
   D = Dir(AppPath & "\*Inventari.SqlTrans")
   While Len(D) > 0
      ReDim Preserve LlistaFiles(UBound(LlistaFiles) + 1)
      LlistaFiles(UBound(LlistaFiles)) = D
      D = Dir
   Wend
On Error GoTo 0
On Error GoTo Noor
   
   For i = 1 To UBound(LlistaFiles)
      Fil = LlistaFiles(i)
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Inventari : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      CalCrearQuery = True
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaInventari(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute
                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaInventari(dataInici) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta) Values(?,?,?,?,?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
'                  Q2.rdoParameters(3) = 0
                  Q2.rdoParameters(4) = ""
                  Q2.rdoParameters(8) = ""
                  SqlExporta1 = "[" & Format(dataFi, "dd-mm-yyyy") & "]"
                  SqlExporta2 = "[" & botiga & "]"
                  DataTot = dataFi
                  CalCrearQuery = False
'                  GuardaHistoric Botiga, DataFi, "Inventari Rebut"
               End If
               Ll = L
               Cm_HInici = Car(L)
               Caca = Car(L)  ' Botiga
               dependenta = Car(L)
               Cm_Article = Car(L)
               Cm_Quantitat = Car(L)
               Proveidor = Car(L) ' TipusDevolucio
               TipusVenta = Car(L)
               Q2.rdoParameters(3) = dependenta
               Q2.rdoParameters(5) = Cm_Article
               Q2.rdoParameters(6) = Cm_Quantitat
               Q2.rdoParameters(7) = ArticleCodiPvp(ArticleCodiGeneticCodi(Cm_Article)) * Cm_Quantitat
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = 0   'Dependenta
               Q2.rdoParameters(4) = Proveidor
               Q2.rdoParameters(8) = TipusVenta
               Q2.Execute
               GuardaHistoric botiga, CVDate(Cm_HInici), "Inventari Rebut", CVDate(Cm_HInici), DependentaCodiNom(dependenta) & " | " & ArticleCodiNom(Cm_Article), CStr(Cm_Quantitat)
            End If
         End If
      
      Wend
      On Error Resume Next
         Q2.Close
         Q1.Close
         CalCrearQuery = True
      On Error GoTo Noor
      TancaFile f
      If Len(SqlExporta1) > 0 Then
         InsertFeineaAFer "TractaInventaris", "DeBotiga", SqlExporta1, SqlExporta2, DataTot
         SqlExporta1 = ""
      End If
      FitcherProcesat Fil
'      If Not TePermisPer("NoAvonarDevolucions") Then ActualizaPanSeco DataInici, CDbl(Botiga)
Mas:
   Next
   Exit Sub

Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub


Sub Interpreta_SqlTrans_Log(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String, D As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String, LlistaFiles() As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, Sa As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As String, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim aa, bb, cc
   
   ReDim LlistaFiles(0)
   D = Dir(AppPath & "\*_log.SqlTrans")
   While Len(D) > 0
      ReDim Preserve LlistaFiles(UBound(LlistaFiles) + 1)
      LlistaFiles(UBound(LlistaFiles)) = D
      D = Dir
   Wend
On Error GoTo 0
On Error GoTo Noor
   
   For i = 1 To UBound(LlistaFiles)
      Fil = LlistaFiles(i)
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Log : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      CalCrearQuery = True
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaLog(dataInici) & "] Where Botiga = ? And TmSt >= ? And TmSt <= ? ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute
                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaLog(dataInici) & "] (Botiga,TmSt,Accio,Params) Values(?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
               End If
               Ll = L
               Cm_HInici = Car(L)
               Cm_Article = Car(L)
               Cm_Otros = Car(L)
               
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = Cm_Article
               Q2.rdoParameters(3) = Cm_Otros
               Q2.Execute
            End If
         End If
      
      Wend
      On Error Resume Next
         Q2.Close
         Q1.Close
         CalCrearQuery = True
      On Error GoTo Noor
      TancaFile f
      FitcherProcesat Fil
Mas:
   Next
   Exit Sub

Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub



Sub Interpreta_SqlTrans_Encarregs(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String, SqlExporta1 As String, SqlExporta2 As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String, D As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String, LlistaFiles() As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double, Caca As String, Sa As String
   Dim Cm_Dependenta As Double, Cm_NumTic As Double, Cm_Article As String, Cm_Quantitat As Double, Cm_Preu As Double, Cm_Import As Double, Cm_Descompte As Double, Cm_Origen As String, Cm_Otros As String, Cm_HInici As Date
   Dim sqlexiste As String
   Dim rsexiste As rdoResultset
   Dim ImportaAuto As Boolean
   
    
    Dim Camp_Id As String, Camp_Dependenta  As String, Camp_Client As String, Camp_data As String, Camp_Estat As String, Camp_Tipus  As String, Camp_Anticip As String, Camp_Botiga As String, Camp_Detall As String, Camp_Periode As String, Camp_Article As String, Camp_Quantitat As String, Camp_Import As String, Camp_Descompte As String, Camp_MarcatDesde As String, Camp_Comentari As String
   
    ReDim LlistaFiles(0)
    D = Dir(AppPath & "\*_Encarregs.SqlTrans")
    While Len(D) > 0
        ReDim Preserve LlistaFiles(UBound(LlistaFiles) + 1)
        LlistaFiles(UBound(LlistaFiles)) = D
        D = Dir
    Wend
    On Error GoTo Noor
   
   For i = 1 To UBound(LlistaFiles)
      Fil = LlistaFiles(i)
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Encarregs : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      CalCrearQuery = True
      While Not EOF(f)
         Line Input #f, L
         If Len(L) > 0 And Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else

               Ll = L
               Camp_Id = Car(L)
               Camp_Dependenta = Car(L)
               Camp_Client = Car(L)
               Camp_data = Car(L)
               Camp_Estat = Car(L)
               Camp_Tipus = Car(L)
               Camp_Anticip = Car(L)
               Camp_Botiga = Car(L)
               Camp_Detall = Car(L)
               Camp_Periode = Car(L)
               Camp_Article = Car(L)
               Camp_Quantitat = Car(L)
               Camp_Import = Car(L)
               Camp_Descompte = Car(L)
               Camp_MarcatDesde = Car(L)
               Camp_Comentari = Car(L)
               
               CreaTaulesDadesTpv CVDate(Camp_data)
               
               Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaencarregs(CVDate(Camp_data)) & "] Where Botiga = ? And Id = ? and Article = ? and (Importat <> 'Si' or Importat is null)")
               Q1.rdoParameters(0) = Camp_Botiga
               'Q1.rdoParameters(1) = Camp_data
               Q1.rdoParameters(1) = Camp_Id
               Q1.rdoParameters(2) = Camp_Article
               Q1.Execute
               
               sqlexiste = "select Id from [" & NomTaulaencarregs(CVDate(Camp_data)) & "] where Id = '" & Camp_Id & "' and Article = '" & Camp_Article & "'"
               Set rsexiste = Db.OpenResultset(sqlexiste)
               If rsexiste.EOF Then
                Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaencarregs(CVDate(Camp_data)) & "] ([Id] , [Dependenta], [Client], [data],[Estat],[Tipus],[Anticip],[Botiga],[Detall],[Periode],[Article],[Quantitat],[Import],[Descompte],[MarcatDesde],[Comentari],[Importat]) Values (? , ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
                Q2.rdoParameters(0) = Camp_Id
                Q2.rdoParameters(1) = Camp_Dependenta
                Q2.rdoParameters(2) = Camp_Client
                Q2.rdoParameters(3) = Camp_data
                Q2.rdoParameters(4) = Camp_Estat
                Q2.rdoParameters(5) = Camp_Tipus
                Q2.rdoParameters(6) = Camp_Anticip
                Q2.rdoParameters(7) = Camp_Botiga
                Q2.rdoParameters(8) = Camp_Detall
                Q2.rdoParameters(9) = Camp_Periode
                Q2.rdoParameters(10) = Camp_Article
                Q2.rdoParameters(11) = Camp_Quantitat
                Q2.rdoParameters(12) = Camp_Import
                Q2.rdoParameters(13) = Camp_Descompte
                Q2.rdoParameters(14) = Camp_MarcatDesde
                Q2.rdoParameters(15) = Camp_Comentari
                Q2.rdoParameters(16) = "No"
                Q2.Execute
                '**********************************************************
                ' IMPORTACIO AUTOMATICA A ALBARANS
                '**********************************************************
                ImportaAuto = False
                Set rsexiste = Db.OpenResultset("select valor from ConstantsEmpresa where camp='ImportaEncarrecs'")
                If Not rsexiste.EOF Then
                    If rsexiste("valor") = "on" Then ImportaAuto = True
                End If
                If ImportaAuto Then
                sqlexiste = "insert into [" & DonamNomTaulaServit(DateValue(Camp_data)) & "] "
                sqlexiste = sqlexiste & "(Client, CodiArticle, PluUtilitzat, Viatge, Equip, QuantitatDemanada, QuantitatServida, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut) "
                sqlexiste = sqlexiste & "select Botiga, article, article, ISNULL(viatge,'Encarrec') viatge, ISNULL(equip,'Encarrec') equip , qd, qs, MotiuModificacio, Hora, TipusComanda, Comentari, ComentariPer, Atribut "
                sqlexiste = sqlexiste & "from "
                sqlexiste = sqlexiste & "(select e.botiga, e.article,  "
                sqlexiste = sqlexiste & "( "
                sqlexiste = sqlexiste & "select viatge from "
                sqlexiste = sqlexiste & "(select top 1 MAX(timestamp) te, max(Viatge) viatge, CodiArticle,client "
                sqlexiste = sqlexiste & "from ComandesMemotecnicPerClient  "
                sqlexiste = sqlexiste & "where (cast(client as varchar)= e.botiga or client is null) and cast(CodiArticle as varchar)=e.Article "
                sqlexiste = sqlexiste & "group by codiArticle, client "
                sqlexiste = sqlexiste & "order by ISNULL(client,0) desc) e1 ) viatge, "
                sqlexiste = sqlexiste & "( "
                sqlexiste = sqlexiste & "select equip from "
                sqlexiste = sqlexiste & "(select top 1 MAX(timestamp) te, max(equip) equip, CodiArticle,client "
                sqlexiste = sqlexiste & "from ComandesMemotecnicPerClient  "
                sqlexiste = sqlexiste & "where (cast(client as varchar)= e.botiga or client is null) and cast(CodiArticle as varchar)=e.Article "
                sqlexiste = sqlexiste & "group by codiArticle, client "
                sqlexiste = sqlexiste & "order by ISNULL(client,0) desc) v1 ) equip, "
                sqlexiste = sqlexiste & "e.quantitat qd, e.quantitat qs, '' AS MotiuModificacio, 69 AS Hora, 2 AS TipusComanda, comentari, '' AS ComentariPer, 0 AS Atribut "
                sqlexiste = sqlexiste & "from [" & NomTaulaencarregs(CVDate(Camp_data)) & "] e "
                sqlexiste = sqlexiste & "where e.estat = 0 and e.id = '" & Camp_Id & "' and e.article = " & Camp_Article & " and isnull(e.importat, 'No')='No' "
                sqlexiste = sqlexiste & ") AA "
                ExecutaComandaSql sqlexiste
                
                'Actualiza encargo
                sqlexiste = "update [" & NomTaulaencarregs(CVDate(Camp_data)) & "] set importat='Si' "
                '**********************************************************
                'SERGI 24/08/2012
                'Update de camp detall posant la mateixa data que el timestamp del servit perque es puguin relacionar
                sqlexiste = sqlexiste & ",detall='[DataCreat:'+left(CONVERT(VARCHAR(10),GETDATE(),105),6)+right(year(GETDATE()),2)+' '+left(CONVERT(VARCHAR(8),GETDATE(),108),5)+']'"
                sqlexiste = sqlexiste & "+substring(detall,27,len(detall)) "
                '**********************************************************
                sqlexiste = sqlexiste & " where estat = 0 and id = '" & Camp_Id & "' and article = " & Camp_Article & " and isnull(importat, 'No')='No' "

                ExecutaComandaSql sqlexiste
                End If
                '**********************************************************
                '**********************************************************
               End If
               
            End If
         End If
      
      Wend
      On Error Resume Next
         Q2.Close
         Q1.Close
         CalCrearQuery = True
      On Error GoTo Noor
      TancaFile f
      FitcherProcesat Fil
Mas:
   Next
   Exit Sub

Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
End Sub




Sub Interpreta_SqlTrans_Salutacions(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double
   Dim Tipus_Moviment As String, import As Double, Motiu As String
   Dim Cm_Data As Date, Cm_Operacio As String, Cm_Dependenta As Double
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*Salutacions.SqlTrans")
   
   On Error GoTo Noor
   
   While Len(Fil) > 0
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Salutacions : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaHoraris(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute
                  InsertFeineaAFer "TraspasHores", "Dependentes", "[" & Format(dataInici, "dd-mm-yyyy ") & " de " & Format(dataInici, "hh:mm") & " a " & Format(dataFi, "hh:mm") & "]", "[" & botiga & "]"

                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaHoraris(dataInici) & "] (Botiga,Data,Dependenta,Operacio) Values(?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
               End If

               Cm_Data = Car(L)
               Cm_Operacio = Car(L)
               Cm_Dependenta = Car(L)
               
               
               Q2.rdoParameters(1) = Cm_Data
               If Not Q2.rdoParameters(1) = Cm_Data Then
                  Q2.rdoParameters(1) = Cm_Data
               End If
               Q2.rdoParameters(2) = Cm_Dependenta
               If Cm_Operacio = "Arriba" Then
                  Q2.rdoParameters(3) = "E"
               ElseIf Cm_Operacio = "Plega" Then
                  Q2.rdoParameters(3) = "P"
               ElseIf Cm_Operacio = "IniciDescans" Then
                  Q2.rdoParameters(3) = "I"
               ElseIf Cm_Operacio = "FiDescans" Then
                  Q2.rdoParameters(3) = "F"
               Else
                  Q2.rdoParameters(3) = "P"
               End If
               Q2.Execute
            End If
         End If
      Wend
      TancaFile f
      FitcherProcesat Fil
Mas:
      Fil = Dir
   Wend
   
   Exit Sub
   
Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas
   
End Sub



Sub Interpreta_SqlTrans_Moviments(Optional Estat As Label = Nothing)
   Dim Contingut As String, NomMaquina As String, nomfile      As String
   Dim Fil As String, f, L As String, Ll As String
   Dim Codi As Double, Preu As Double, nom As String
   Dim Q As rdoQuery, i As Integer, Q1 As rdoQuery, Q2 As rdoQuery
   Dim LlistaCamps As String, NomTaula As String
   Dim CalCrearQuery As Boolean, dataInici As Date, botiga As Integer, dataFi As Date, Di As Date, Df As Date, Cb As Integer
   Dim data As Date, dependenta As Double
   Dim Tipus_Moviment As String, import As Double, Motiu As String
   Dim Cm_Primera As String, Cm_Oberta As String, Caca As String, Cm_Tancada As String
   Dim Cm_CanviInicial As String, Cm_CanviInicialDetall  As String, Cm_CanviFinal As String, Cm_CanviFinalDetall As String, Cm_PrimerTicket As String, Cm_UltimTicket As String
   Dim Cm_HInici As Date, Cm_HFi As Date, Cm_Zeta As String, Cm_descuadre As String, Cm_Clients As String
   Dim Ca As String, P As Integer
   Dim Cm_Import As Double, CM_Motiu As String
   Dim Cm_NumZeta As Integer
   Dim Q3 As rdoQuery
   
   CalCrearQuery = True
   Fil = Dir(AppPath & "\*Moviments.SqlTrans")
   
   On Error GoTo Noor
   
   While Len(Fil) > 0
      f = FreeFile
      If Not Estat Is Nothing Then Estat.Caption = "Interpretant Moviments : " & Fil
      My_DoEvents
      Open AppPath & "\" & Fil For Input As #f
      DescomposaContingut Fil, Contingut, NomMaquina, nomfile
      While Not EOF(f)
         Line Input #f, L
         If Not Left(L, 1) = "#" Then
            If Left(L, 5) = "[Sql-" Then
               Ll = DonamParam(L)
               If Left(L, 15) = "[Sql-DataInici:" Then
                  CalCrearQuery = True
                  dataInici = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-DataFi:" Then
                  CalCrearQuery = True
                  dataFi = CVDate(Ll)
               End If
               If Left(L, 12) = "[Sql-Botiga:" Then
                  CalCrearQuery = True
                  botiga = Ll
                  LogaHw 1, CDbl(NomMaquina), Ll, "La Llicencia : " & NomMaquina & " Correspon a la botiga " & Ll
               End If
            Else
               If CalCrearQuery Then
                  CreaTaulesDadesTpv dataInici
                  Set Q1 = Db.CreateQuery("", "Delete [" & NomTaulaMovi(dataInici) & "] Where Botiga = ? And Data >= ? And Data <= ? And (Tipus_Moviment='O' Or Tipus_Moviment='A') ")
                  Q1.rdoParameters(0) = botiga
                  Q1.rdoParameters(1) = dataInici
                  Q1.rdoParameters(2) = dataFi
                  Q1.Execute

                  Set Q2 = Db.CreateQuery("", "Insert Into [" & NomTaulaMovi(dataInici) & "] (Botiga,Data,DependentA,Tipus_Moviment,Import,Motiu) Values (?,?,?,?,?,?) ")
                  Q2.rdoParameters(0) = botiga
                  CalCrearQuery = False
               End If
               Cm_HInici = Car(L)
               Cm_NumZeta = Car(L)
               Cm_Primera = Car(L)
               Cm_Import = Car(L)
               CM_Motiu = Car(L)
        
               Q2.rdoParameters(1) = Cm_HInici
               Q2.rdoParameters(2) = Cm_Primera
               
               If Cm_Import > 0 Then
                  Q2.rdoParameters(3) = "A"
               Else
                  Q2.rdoParameters(3) = "O"
               End If
               
               Q2.rdoParameters(4) = Cm_Import
               Q2.rdoParameters(5) = CM_Motiu
               Q2.Execute
            End If
         End If
      Wend
      TancaFile f
      FitcherProcesat Fil
      
        'Sincronitzacio MySQL Ametller Arqueigs
        'If UCase(EmpresaActual) = UCase("LaForneria") Then
        '  InsertFeineaAFer "SincroDbArqueigsAmetller", "[" & botiga & "]", "[" & DataInici & "]", "", "[" & dataFi & "]"
        'End If
     
        'Sincronitzacio MURANO
        'If UCase(EmpresaActual) = UCase("Tena") And (botiga = "384" Or botiga = "214") Then 'SOLO T--14, T--35
        '  InsertFeineaAFer "SincroMURANOMoviments", "[" & botiga & "]", "[" & DataInici & "]", "[" & dataFi & "]", nomFile
        'End If
     
Mas:
        Fil = Dir
    Wend


   Exit Sub
   
Noor:
   TancaFile f
   FitcherProcesat Fil, False
   Resume Mas

End Sub






Sub FesFeina()
   Dim i As Integer, Fil As String, Resultat, Dir_Desti As String
   Dim D1 As Date, D2 As Date, LlistaFiles() As String, D As String
   Dim Estat As String
   
   DIRECTORI_TEMPORAL = AppPath & "\"
   DIRECTORI_TEMPORAL_BAK = AppPath & "\Bak\"
   DIRECTORI_TEMPORAL_ERR = AppPath & "\Err\"
   
   ReDim LlistaFiles(0)
   D = Dir(DIRECTORI_TEMPORAL & "*??-??-??.*")
   While Len(D) > 0
      If Not Right(D, 1) = "e" Then
         ReDim Preserve LlistaFiles(UBound(LlistaFiles) + 1)
         LlistaFiles(UBound(LlistaFiles)) = D
      End If
      D = Dir
   Wend
   
   
   For i = 1 To UBound(LlistaFiles)
      Fil = LlistaFiles(i)
      D1 = Now
      
      Resultat = Converteix(Fil)
      
      On Error Resume Next
      
      If Resultat = 1 Then
         Dir_Desti = DIRECTORI_TEMPORAL_BAK
         Estat = "Ok "
      Else
         Dir_Desti = DIRECTORI_TEMPORAL_ERR
         Estat = "Err"
      End If
      
      Name DIRECTORI_TEMPORAL & Fil As Dir_Desti & Fil
'      Kaill DIRECTORI_TEMPORAL & Fil
      Name DIRECTORI_TEMPORAL & Fil & "e" As Dir_Desti & Fil & "e"
'      Kaill DIRECTORI_TEMPORAL & Fil & "e"
      
      D2 = Now
   Next
   
End Sub



Function ActualizaPanSeco(D1 As Date, Cl As Double) As Boolean
   Dim D As Date
   Dim sql As String, Memo As String, Viatge As String, equip As String, SQL1 As String
   Dim Cantidad As Double, Plu As Double, codiArticle As Double
   Dim rs As rdoResultset, rs1 As rdoResultset, Rs2 As rdoResultset
   Dim sql2 As String, sql3 As String, sql4 As String
   Dim QCodiArt As rdoQuery, Q_BorraTots As rdoQuery, CreaNou As rdoQuery, QActualitza As rdoQuery
      
   If Not ControlEstat Is Nothing Then ControlEstat.Caption = FilActual & "  Actualitzant Pa Sec "
   My_DoEvents
      
On Error GoTo nor
      
      D = D1
      ActualizaPanSeco = False
      
      DonamNomTaulaServit D
      Set rs = Db.OpenResultset("Select * From [" & NomTaulaDevol(D) & "] Where Day(Data) = " & Day(D) & " And Botiga = " & Cl & " ")
      
      sql = "Insert Into [Servit-" & Format(D, "yy-mm-dd") & "] "
      sql = sql & "(Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut) "
      sql = sql & "Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
      Set CreaNou = Db.CreateQuery("", sql)
      ExecutaComandaSql "Update [Servit-" & Format(D, "yy-mm-dd") & "] Set QuantitatTornada = 0 Where Client = " & Cl & " And Viatge = 'Auto' And Equip = 'Auto' And QuantitatDemanada = 0 And QuantitatServida = 0 "
      
      While Not rs.EOF
         Plu = rs("Plu")
         Cantidad = rs("Quantitat")
         codiArticle = ArticleCodiGeneticCodi(Plu)
         If codiArticle = 0 Then
            codiArticle = codiArticle
         End If
         CreaNou.rdoParameters(0) = Cl          'Client
         CreaNou.rdoParameters(1) = codiArticle 'CodiArticle
         CreaNou.rdoParameters(2) = Plu         'PluUtilitzat
         CreaNou.rdoParameters(3) = "Auto"      'Viatge
         CreaNou.rdoParameters(4) = "Auto"      'Equip
         CreaNou.rdoParameters(5) = 0           'QuantitatDemanada
         CreaNou.rdoParameters(6) = Cantidad    'QuantitatTornada
         CreaNou.rdoParameters(7) = 0           'QuantitatServida
         CreaNou.rdoParameters(8) = ""          'MotiuModificacio
         CreaNou.rdoParameters(9) = 0           'Hora
         CreaNou.rdoParameters(10) = 2          'TipusComanda
         CreaNou.rdoParameters(11) = ""         'Comentari
         CreaNou.rdoParameters(12) = ""         'ComentariPer
         CreaNou.rdoParameters(13) = 0          'Atribut
         CreaNou.Execute
         
         ActualizaPanSeco = True
         rs.MoveNext
      Wend
      rs.Close

nor:

End Function
Sub BuscaAlgunMemotecnic(CD As Double, Memo As String, Viatge As String, equip As String)
   Dim rs As rdoResultset
   
   Memo = ""
   Viatge = ""
   equip = ""
   
   Set rs = Db.OpenResultset("Select Memotecnic,Viatge,Equip  From Memotecnics Where  Codi = " & CD & " ")
   If Not rs.EOF Then
      Memo = rs("Memotecnic")
      Viatge = rs("Viatge")
      equip = rs("Equip")
   End If
   rs.Close
   
End Sub



Sub CarregaHashPreus()
   Dim rs As rdoResultset, i As Integer
   
   For i = 0 To UBound(HashPreus)
      HashPreus(i) = 0
   Next
   If Not ExisteixTaula("Articles") Then Exit Sub
   
   Set rs = Db.OpenResultset("Select CodiGenetic,isnull(Preu,0) Preu From Articles with (nolock)")
   While Not rs.EOF
      If rs("CodiGenetic") > 0 And rs("CodiGenetic") < UBound(HashPreus) Then HashPreus(rs("CodiGenetic")) = rs("Preu")
      rs.MoveNext
   Wend
   rs.Close
   
'
'   PreuProducte.rdoParameters(0) = Plu
'   Set Rs = PreuProducte.OpenResultset
'   Article_Plu_Preu = 0
'   If Not Rs.EOF Then If Not IsNull(Rs("Preu")) Then Article_Plu_Preu = Rs("Preu")
'   Rs.Close

End Sub

Function ColorDeNumARGB(i As Integer)
   Dim Col As Integer
   
   Select Case i
      Case 1: Col = 8
      Case 2: Col = 1
      Case 3: Col = 11
      Case 4: Col = 10
      Case 5: Col = 12
      Case 6: Col = 5
      Case 7: Col = 14
      Case 8: Col = 15
      Case 9: Col = 3
      Case 10: Col = 9
      Case 11: Col = 2
      Case 12: Col = 13
      Case 13: Col = 4
      Case 14: Col = 6
      Case 15: Col = 7
   End Select
   
   ColorDeNumARGB = QBColor(Col)
   
End Function




Sub CreaTaulesDadesTpv(D As Date)
   Dim sql As String
    
   sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25), Plu float, Quantitat float, Import float,Tipus_venta [nvarchar] (25),FormaMarcar [nvarchar] (255) Default (''),Otros [nvarchar] (255) Default ('')"
   If Not ExisteixTaula("[" & NomTaulaVentas(D) & "]") Then
      ExecutaComandaSql "Create Table [" & NomTaulaVentas(D) & "] (" & sql & ") "
      ExecutaComandaSql "CREATE NONCLUSTERED INDEX [" & NomTaulaVentas(D) & "] ON [dbo].[" & NomTaulaVentas(D) & "]([data])"
   End If
   
   If Not ExisteixTaula("[" & NomTaulaAlbarans(D) & "]") Then
      ExecutaComandaSql "Create Table [" & NomTaulaAlbarans(D) & "] (" & sql & ") "
   End If
   
   ' Com Ventas  Sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25), Plu float, Quantitat float, Import float,Tipus_venta [nvarchar] (25) "
   If Not ExisteixTaula("[" & NomTaulaAnulats(D) & "]") Then
      ExecutaComandaSql "Create Table [" & NomTaulaAnulats(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaDevol(D) & "]") Then
      sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25),Plu float,Quantitat float,Import float,Tipus_venta [nvarchar] (25) "
      ExecutaComandaSql "Create Table [" & NomTaulaDevol(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaLog(D) & "]") Then
      sql = "TmSt datetime,Botiga float,Accio [nvarchar] (255),Params [nvarchar] (255) "
      ExecutaComandaSql "Create Table [" & NomTaulaLog(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaInventari(D) & "]") Then
      sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25),Plu float,Quantitat float,Import float,Tipus_venta [nvarchar] (25) "
      ExecutaComandaSql "Create Table [" & NomTaulaInventari(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaMovi(D) & "]") Then
      sql = "Botiga float,Data datetime,Dependenta float,Tipus_moviment [nvarchar] (25),Import float, Motiu [nvarchar] (25)"
      ExecutaComandaSql "Create Table [" & NomTaulaMovi(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaHoraris(D) & "]") Then
      sql = "Botiga float,Data datetime,Dependenta float,Operacio [nvarchar] (25)"
      ExecutaComandaSql "Create Table [" & NomTaulaHoraris(D) & "] (" & sql & ") "
   End If
   
   If Not ExisteixTaula("[" & NomTaulaencarregs(D) & "]") Then
      sql = "[Id][nvarchar] (255),[Dependenta][float],[Client][nvarchar] (255),[Data][DateTime],[Estat][float],[Tipus][float],[Anticip][float],[Botiga][float],[Detall][nvarchar] (255),[Periode][nvarchar] (255),[Article][float],[Quantitat][float],[Import][float],[Descompte][nvarchar] (255),[MarcatDesde][nvarchar] (255),[Comentari][nvarchar] (255),[Importat][nvarchar] (255)"
      ExecutaComandaSql "Create Table [" & NomTaulaencarregs(D) & "] (" & sql & ") "
   End If
      
   AseguraExisteixTaulaAlertes D
   AseguraExisteixTaulaPrevisions D
      
End Sub



Function FormatHora(MinTot As Double)
   Dim Min As Integer, Hor As Integer
   
   
   Hor = Int(MinTot / 60)
   Min = MinTot - (Hor * 60)
   FormatHora = Format(Hor, "00") & ":" & Format(Min, "00")
End Function


Sub esborra_errors(fili)
   Dim st_base As String
   Dim cadena As String
   Dim NumZeta As Integer
   
   NumZeta = 1
   If Len(Dir(fili & "e")) > 0 Then
      num_tickets_esborrats = 0
      ReDim tickets_esborrats(num_tickets_esborrats)
     
      Ultima_Zeta = DateValue(data_en_conversio) + TimeValue("00:00:00")
      Pen_Ultima_Zeta = Ultima_Zeta
      
      Handel2 = FreeFile
      Open fili & "e" For Input As Handel2
      dos_opened = True
      My_DoEvents
         If Not EOF(Handel2) Then Line Input #Handel2, cadena ' Treiem pos de la seg zeta
         While Not EOF(Handel2)
            Line Input #Handel2, cadena
            Select Case Left(cadena, 1)
               Case "Z"
                 tracta_errors_z cadena, NumZeta
               Case "M"
                 tracta_errors_ticket cadena
            End Select
         Wend
   End If
End Sub


Function file_data(nom_file)
   Dim pos As Integer
   pos = InStr(nom_file, ".")
   file_data = CDate(Mid(nom_file, 1, pos - 1))
End Function

Function File_CodiBotiga(nom_file)
    Dim P As Integer, Codi As Integer
    
    P = InStr(nom_file, ".")
    Codi = Val(Mid(nom_file, P + 1, 2))
    
'    file_nombotiga = botiga_codi_nom(Codi)
    File_CodiBotiga = Codi
    
End Function
Sub InitVectorDependentes()
   ReDim DependentesALaBotiga(0)
   NumDependentesALaBotiga = 0
End Sub

Sub LogaHw(Tipus As Double, Codi As Double, Valor1 As String, Texte As String, Optional Valor2 As String = "", Optional Valor3 As String = "")
   Dim rs As rdoResultset
   Dim Rs2 As rdoResultset
   Dim Rs3 As rdoResultset
   Dim Rs4 As rdoResultset
   Dim Rs5 As rdoResultset
   Dim Rs6 As rdoResultset
   Dim rstest As Recordset
   Dim cliente1 As String
   Dim sqlrecurso As String
   Dim nuevoid As String
   Dim prueba
   If Not ExisteixTaula("ParamsHw") Then ExecutaComandaSql "CREATE TABLE [ParamsHw] ([Tipus] [numeric](18, 0) NULL ,[Codi] [numeric](18, 0) NULL ,[Valor1] [nvarchar] (255) NULL ,[Valor2] [nvarchar] (255) NULL ,[Valor3] [nvarchar] (255) NULL ,[Valor4] [nvarchar] (255) NULL ,[Descripcio] [nvarchar] (255) NULL) ON [PRIMARY]"
   
   ExecutaComandaSql "Delete ParamsHw Where  Tipus = " & Tipus & " And Codi = " & Codi & " "
   ExecutaComandaSql "Insert Into ParamsHw  ([Tipus],[Codi],[Valor1],[Valor2],[Valor3],[Valor4],[Descripcio]) Values (" & Tipus & "," & Codi & ",'" & Valor1 & "','" & Valor2 & "','" & Valor3 & "','','" & Texte & "') "
   
   Set rs = Db.OpenResultset("Select * from clients where codi = '" & Valor1 & "'")
   If rs.EOF Then
      ExecutaComandaSql "Insert into Clients (codi,nom,nif,adresa ,[Desconte 5]) values (" & Valor1 & ",'Botiga Llicencia " & Valor1 & "','','',0)"
   End If
    
   Set rs = Db.OpenResultset("Select * from hit.dbo.llicencies where llicencia = '" & Codi & "' ")
   If rs.EOF Then
        ExecutaComandaSql "Delete hit.dbo.llicencies where llicencia = '" & Codi & "'"
        ExecutaComandaSql "Insert into hit.dbo.llicencies values (" & Codi & ",'" & EmpresaActual & "','Desde Sincro llicencia nova','','',GETDATE())"
            sf_enviarMail "Secrehit@gmail.com", EmailGuardia, "NUEVA LICENCIA CREADA", "La llicencia " & Codi & " se ha creado para el cliente " & EmpresaActual, "", ""
   Else
       cliente1 = rs("Empresa")
       Set rs = Db.OpenResultset("Select * from hit.dbo.llicencies where llicencia = '" & Codi & "' and upper(empresa) = '" & UCase(EmpresaActual) & "'")
       If rs.EOF Then
          ExecutaComandaSql "Delete hit.dbo.llicencies where llicencia = '" & Codi & "'"
          ExecutaComandaSql "Insert into hit.dbo.llicencies values (" & Codi & ",'" & EmpresaActual & "','Desde Sincro llicencia nova')"
          sf_enviarMail "Secrehit@gmail.com", EmailGuardia, "LICENCIA CAMBIADA DE CLIENTE", "La llicencia " & Codi & " ha cambiado del cliente " & cliente1 & " al cliente " & EmpresaActual, "", ""
       End If
   End If
   
    'sqlrecurso = "select rec.id from Fac_JordiBosch.dbo.recursos rec with (nolock)"
    'sqlrecurso = sqlrecurso & " left join Fac_JordiBosch.dbo.recursosExtes recex with (nolock)"
    'sqlrecurso = sqlrecurso & " on rec.Id = recex.Id and Variable = 'LICENCIA'"
    'sqlrecurso = sqlrecurso & "where Tipo = 'LLOC' and Valor = '" & Codi & "'"
    'Set rs = Db.OpenResultset(sqlrecurso)
    'If rs.EOF Then
    '      'sqlrecurso = "select * from clients where codi = " & Valor1
    '      sqlrecurso = " select * from clients left join constantsclient on clients.codi = ConstantsClient.codi"
    '      sqlrecurso = sqlrecurso & " where clients.codi = '" & Valor1 & "' and Variable = 'Tel'"
    '      Set Rs2 = Db.OpenResultset(sqlrecurso)
    '      If Rs2.EOF Then
'   '         sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "TIENDA SIN RECURSO Y SIN CLIENTE ASOCIADO", "La llicencia " & Codi & " del cliente " & EmpresaActual & " no tiene recurso asignado y no existe cliente en la base de datos del cliente", "", ""
    '      Else
    '           sqlrecurso = "select newid()"
    '           Set rs3 = Db.OpenResultset(sqlrecurso)
    '           If Not rs3.EOF Then
    '                nuevoid = rs3(0)
    '                ExecutaComandaSql "insert into fac_jordibosch.dbo.recursos values ('" & nuevoid & "','" & Rs2("nom") & "','LLOC')"
    '                'ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "','LICENCIA','" & Codi & "')"
    '                sqlrecurso = " select replace(replace(Valor,'-',''),' ','') from ConstantsEmpresa"
    '                sqlrecurso = sqlrecurso & " where Camp like '%CampNif'"
    '                sqlrecurso = sqlrecurso & " and substring(Camp,0,charindex('_',Camp,0)) = ("
    '                sqlrecurso = sqlrecurso & " select substring(Camp,0,charindex('_',Camp,0)) from ConstantsEmpresa"
    '                sqlrecurso = sqlrecurso & " where Camp like '%CampCliente%'"
    '                sqlrecurso = sqlrecurso & " and Valor = '" & Valor1 & "')"
    '                Set Rs4 = Db.OpenResultset(sqlrecurso)
    '                If Not Rs4.EOF Then
    '                    sqlrecurso = " select codi from Fac_JordiBosch.dbo.clients"
    '                    sqlrecurso = sqlrecurso & " where replace(replace(nif,'-',''),' ','') like '%" & Rs4(0) & "%'"
    '                    Set Rs5 = Db.OpenResultset(sqlrecurso)
    '                    If Not Rs5.EOF Then
    '                        ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "', 'CLIENTE', '" & Rs5(0) & "')"
    '                    Else
'   '                         sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "AL NUEVO RECURSO NO SE LE PUDO ASOCIAR RAZON SOCIAL", "La llicencia " & Codi & " del cliente " & EmpresaActual & " se le generó recurso pero no se le pudo asignar razon social.", "", ""
    '                    End If
    '                Else
    '                'isnull(replace(isnull(Valor,''),'-',''),'')
    '                    sqlrecurso = "select Valor numero  from dbo.ConstantsEmpresa where Camp like '%CampNif' and substring(Camp,0,charindex('_',Camp,0))=''"
    '                    Set Rs5 = Db.OpenResultset(sqlrecurso)
    '                    '******************************************************
    '                    ' JORGE MODIFICADO 16/07/2011 linea siguiente estaba mal hay que comprobar primero eof y luego
    '                    ' por si acaso en rs5(0) sea nulo, pq si rs5 eof no se puede mirar rs5(0)
    '                    '******************************************************
    '                    If Not Rs5.EOF Then
    '                        If Rs5(0) <> "Nulo" Then
    '                            prueba = Rs5(0)
    '                            prueba = Replace(prueba, "-", "")
    '                            sqlrecurso = " select codi from Fac_JordiBosch.dbo.clients"
    '                            sqlrecurso = sqlrecurso & " where replace(replace(nif,'-',''),' ','') like '%" & Rs5(0) & "%'"
    '                            Set Rs6 = Db.OpenResultset(sqlrecurso)
    '                            If Not Rs6.EOF Then
    '                                ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "', 'CLIENTE', '" & Rs6(0) & "')"
    '                            Else
'   '                                 sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "AL NUEVO RECURSO NO SE LE PUDO ASOCIAR RAZON SOCIAL", "La llicencia " & Codi & " del cliente " & EmpresaActual & " se le generó recurso pero no se le pudo asignar razon social.", "", ""
    '                            End If
    '                            Rs6.Close
    '                        Else
'   '                             sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "AL NUEVO RECURSO NO SE LE PUDO ASOCIAR RAZON SOCIAL", "La llicencia " & Codi & " del cliente " & EmpresaActual & " se le generó recurso pero no se le pudo asignar razon social.", "", ""
    '                        End If
    '                    Else
'   '                         sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "AL NUEVO RECURSO NO SE LE PUDO ASOCIAR RAZON SOCIAL", "La llicencia " & Codi & " del cliente " & EmpresaActual & " se le generó recurso pero no se le pudo asignar razon social.", "", ""
    '                    End If
    '                    Rs5.Close
    '                End If
    '                Rs4.Close
    '                sqlrecurso = "insert into fac_jordibosch.dbo.recursosExtes values ('" & nuevoid & "','EMPRESA_HIT',"
    '                sqlrecurso = "(select id from hit.dbo.web_empreses where nom like '" & EmpresaActual & "'))"
    '                ExecutaComandaSql sqlrecurso
    '                ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "', 'LICENCIA', '" & Codi & "')"
    '                ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "', 'DIRECCION', '" & Rs2("adresa") & " " & Rs2("Ciutat") & "')"
    '                ExecutaComandaSql "insert into fac_jordibosch.dbo.recursosextes values ('" & nuevoid & "', 'TELEFONOS', '" & Rs2("adresa") & " " & Rs2("Valor") & "')"
'   '                 sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "NUEVO RECURSO", "La llicencia " & Codi & " de la tienda " & Rs2("nom") & " del cliente " & EmpresaActual & " se le ha generado un nuevo recurso.", "", ""
    '            End If
    '            rs3.Close
    '      End If
    '      Rs2.Close
    'End If
    'rs.Close
End Sub

Sub LogaErr(Str As String)
   
   If Not ExisteixTaula("ErrorsSincro") Then ExecutaComandaSql "CREATE TABLE [ErrorsSincro] ([Valor1] [nvarchar] (255) NULL ,[Valor2] [nvarchar] (255) NULL ,[Valor3] [nvarchar] (255) NULL ,[Valor4] [nvarchar] (255) NULL ,[Descripcio] [nvarchar] (255) NULL) ON [PRIMARY]"
 
   
   ExecutaComandaSql "Insert Into ErrorsSincro  ([Valor1],[Valor2]) Values ('" & Now() & "','" & Str & "') "

End Sub

Sub marca_ultim_ticket_com_a_impres()
   
   u_t_estat = "Impres"

End Sub

Function nom_database(Codi_Botiga, data)
  nom_database = Format(Val(Codi_Botiga)) & "-" & Format(data, "mm") & "-" & Format(data, "yy")
End Function


Function NomFileCurt(Llarg As String) As String
   Dim P As Integer, Pp As Integer
   
   Pp = 0
   Do
      P = InStr(Pp + 1, Llarg, "]")
      If P > 0 Then Pp = P
   Loop While P > 0
   
   NomFileCurt = Llarg
   If Pp > 0 Then
      NomFileCurt = Right(Llarg, Len(Llarg) - Pp)
   End If
   
End Function

Function DonamNomTaulaServit(dia As Date) As String
    DonamNomTaulaServit = "Servit-" & Format(dia, "yy-mm-dd")
    If Not ExisteixTaula(DonamNomTaulaServit) Then CreaTaulaServit DonamNomTaulaServit
    If Not ExisteixTaula(DonamNomTaulaServit & "Trace") Then CreaTaulaServitArregla DonamNomTaulaServit
End Function

Function DonamNomTaulaEstimacio(dia As Date) As String
    DonamNomTaulaEstimacio = "Estimat-" & Format(dia, "yy-mm-dd")
    If Not ExisteixTaula(DonamNomTaulaEstimacio) Then CreaTaulaEstimat DonamNomTaulaEstimacio
    
End Function



Function nomTaulaOracul(data As Date)
    
'    nomTaulaOracul = "V_Oracul_" & Format(Data, "yyyy-mm")
    
'    If Not ExisteixTaula(nomTaulaOracul) Then
'        sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25), Plu float, Quantitat float, Import float,Tipus_venta [nvarchar] (25),FormaMarcar [nvarchar] (255) Default (''),Otros [nvarchar] (255) Default ('')"
'        ExecutaComandaSql "Create Table [" & nomTaulaOracul & "] (" & sql & ") "
'    End If

End Function

Function tablaAnaliticaSemanal(anyo As Integer, Semana As Integer) As String
    Dim sql As String

    tablaAnaliticaSemanal = "ANALITICA_PRODUCTOS_" & anyo & "_" & Semana
    If Not ExisteixTaula(tablaAnaliticaSemanal) Then
        sql = "CREATE TABLE " & tablaAnaliticaSemanal & " ( "
        sql = sql & "[Emp_HIT] [int] NULL, "
        sql = sql & "[Emp_SAGE] [int] NULL, "
        sql = sql & "[Ejercicio] [int] NULL, "
        sql = sql & "[Asiento_SAGE] [numeric](10, 0) NULL, "
        sql = sql & "[Semana] [int] NULL, "
        sql = sql & "[IdFactura] [nvarchar](255) NULL, "
        sql = sql & "[FechaFactura] [datetime] NULL, "
        sql = sql & "[NumFactura] [nvarchar](255) NULL, "
        sql = sql & "[ClienteProveedorCodi] [nvarchar](255) NULL, "
        sql = sql & "[ClienteProveedor] [nvarchar](255) NULL, "
        sql = sql & "[TipoFactura] [nvarchar](5) NULL, "
        sql = sql & "[TablaFactura] [nvarchar](100) NULL, "
        sql = sql & "[Producto] [numeric](18, 0) NULL, "
        sql = sql & "[Familia] [nvarchar](255) NULL, "
        sql = sql & "[Importe] [float] NULL, "
        sql = sql & "[Msg_Importacion] [nvarchar](255) NULL, "
        sql = sql & "[MP] [nvarchar](255) NULL"
        sql = sql & "   )"
        ExecutaComandaSql sql
    End If

End Function

Function taulaCodigosDeAccion() As String
    Dim sql As String
    
    If Not ExisteixTaula("CodigosDeAccion") Then
        sql = "CREATE TABLE [dbo].[CodigosDeAccion]("
        sql = sql & "[IdCodigo] [nvarchar](255) NOT NULL,"
        sql = sql & "[TmStmp] [datetime] NOT NULL,"
        sql = sql & "[TipoCodigo] [nvarchar](255) NOT NULL,"
        sql = sql & "[Param1] [nvarchar](255) NULL,"
        sql = sql & "[Param2] [nvarchar](255) NULL,"
        sql = sql & "[Param3] [nvarchar](255) NULL,"
        sql = sql & "[Param4] [nvarchar](255) NULL,"
        sql = sql & "[Param5] [nvarchar](255) NULL,"
        sql = sql & "[Param6] [nvarchar](255) NULL,"
        sql = sql & "[Param7] [nvarchar](255) NULL,"
        sql = sql & "[Param8] [nvarchar](255) NULL,"
        sql = sql & "[Param9] [nvarchar](255) NULL,"
        sql = sql & "[Param10] [nvarchar](255) NULL"
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
    End If
    taulaCodigosDeAccion = "CodigosDeAccion"
    
End Function

Function taulaServits(ByVal D) As String
    Dim NomTaula As String
    Dim SqlCTS As String, SqlTrace As String, SqlTrigger As String
    Dim TaulaServitsTrace As String
    Dim Y As Integer
    
    Y = Year(D)
    If Year(D) > Year(Now) + 1 Then Y = Year(Now) 'Si estan pidiendo movimientos de un año futuro NO se crea la tabla
    If Year(D) < 1998 Then Y = 1998               'Si estan pidiendo movimientos de un año anterior a 1998 NO se crea la tabla

    taulaServits = "[Servit-" & Right(Y, 2) & "-" & Right("00" & Month(D), 2) & "-" & Right("00" & Day(D), 2) & "]"
    TaulaServitsTrace = "[Servit-" & Right(Y, 2) & "-" & Right("00" & Month(D), 2) & "-" & Right("00" & Day(D), 2) & "Trace]"
    NomTaula = "Servit-" & Right(Y, 2) & "-" & Right("00" & Month(D), 2) & "-" & Right("00" & Day(D), 2) & ""
    
    SqlCTS = "CREATE TABLE " & taulaServits & " ("
    SqlCTS = SqlCTS & "[Id]                [uniqueidentifier] Default (NEWID()),"
    SqlCTS = SqlCTS & "[TimeStamp]         [datetime]              Default (GetDate())  ,"
    SqlCTS = SqlCTS & "[QuiStamp]          [nvarchar] (255)        Default (Host_Name()),"
    SqlCTS = SqlCTS & "[Client]            [float]          Null ,"
    SqlCTS = SqlCTS & "[CodiArticle]       [int]            Null ,"
    SqlCTS = SqlCTS & "[PluUtilitzat]      [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[Viatge]            [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[Equip]             [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[QuantitatDemanada] [float]                 Default (0),"
    SqlCTS = SqlCTS & "[QuantitatTornada]  [float]                 Default (0),"
    SqlCTS = SqlCTS & "[QuantitatServida]  [float]                 Default (0),"
    SqlCTS = SqlCTS & "[MotiuModificacio]  [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[Hora]              [float]          Null ,"
    SqlCTS = SqlCTS & "[TipusComanda]      [float]          Null ,"
    SqlCTS = SqlCTS & "[Comentari]         [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[ComentariPer]      [nvarchar] (255) Null ,"
    SqlCTS = SqlCTS & "[Atribut]           [Int]            Null ,"
    SqlCTS = SqlCTS & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
    SqlCTS = SqlCTS & "[CitaServida]       [nvarchar] (255)        Default (''),"
    SqlCTS = SqlCTS & "[CitaTornada]       [nvarchar] (255)        Default ('') "
    SqlCTS = SqlCTS & ") ON [PRIMARY]"
    
    SqlTrace = "CREATE TABLE " & TaulaServitsTrace & " ("
    SqlTrace = SqlTrace & "[Modificat]         [datetime]              Default (GetDate())  ,"
    SqlTrace = SqlTrace & "[Id]                [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[TimeStamp]         [datetime]       Null ,"
    SqlTrace = SqlTrace & "[QuiStamp]          [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[Client]            [float]          Null ,"
    SqlTrace = SqlTrace & "[CodiArticle]       [int]            Null ,"
    SqlTrace = SqlTrace & "[PluUtilitzat]      [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[Viatge]            [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[Equip]             [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[QuantitatDemanada] [float]                 Default (0),"
    SqlTrace = SqlTrace & "[QuantitatTornada]  [float]                 Default (0),"
    SqlTrace = SqlTrace & "[QuantitatServida]  [float]                 Default (0),"
    SqlTrace = SqlTrace & "[MotiuModificacio]  [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[Hora]              [float]          Null ,"
    SqlTrace = SqlTrace & "[TipusComanda]      [float]          Null ,"
    SqlTrace = SqlTrace & "[Comentari]         [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[ComentariPer]      [nvarchar] (255) Null ,"
    SqlTrace = SqlTrace & "[Atribut]           [Int]            Null ,"
    SqlTrace = SqlTrace & "[CitaDemanada]      [nvarchar] (255)        Default (''),"
    SqlTrace = SqlTrace & "[CitaServida]       [nvarchar] (255)        Default (''),"
    SqlTrace = SqlTrace & "[CitaTornada]       [nvarchar] (255)        Default ('') "
    SqlTrace = SqlTrace & ") ON [PRIMARY]"
    
    SqlTrigger = "CREATE TRIGGER [M_" & NomTaula & "] ON [" & NomTaula & "] "
    SqlTrigger = SqlTrigger & "AFTER INSERT,UPDATE,DELETE AS "
    SqlTrigger = SqlTrigger & "Update [" & NomTaula & "] Set [TimeStamp] = GetDate(), [QuiStamp] = Host_Name() Where Id In (Select Id From Inserted) "
    SqlTrigger = SqlTrigger & "Insert Into ComandesModificades Select Id As Id,GetDate() As [TimeStamp],'" & NomTaula & "' As TaulaOrigen From Inserted "
    SqlTrigger = SqlTrigger & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from [" & NomTaula & "] Where Id In (Select Id From Inserted)"
    SqlTrigger = SqlTrigger & "Insert into [" & NomTaula & "Trace] (Id,[TimeStamp],[QuiStamp],Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada) Select Id,[TimeStamp],[QuiStamp]+'BORRAT!!!',Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada,QuantitatTornada,QuantitatServida,MotiuModificacio,Hora,TipusComanda,Comentari,ComentariPer,Atribut,CitaDemanada  from deleted Where not Id In (Select Id From Inserted) "
    SqlTrigger = SqlTrigger & "If Update (QuantitatTornada)  Update [" & NomTaula & "] Set [CitaTornada]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatTornada  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaTornada]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
    SqlTrigger = SqlTrigger & "If Update (QuantitatServida)  Update [" & NomTaula & "] Set [CitaServida]  = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatServida  AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaServida]  AS VarChar(255)) Where Id In (Select Id From Inserted) "
    SqlTrigger = SqlTrigger & "If Update (QuantitatDemanada) Update [" & NomTaula & "] Set [CitaDemanada] = Cast(CAST(Host_Name() AS VarChar(255)) + ',' + CAST(QuantitatDemanada AS VarChar(255))  + ',' + CAST(GetDate() AS VarChar(255))  +  '/' + [CitaDemanada] AS VarChar(255)) Where Id In (Select Id From Inserted) "
    
    If Not ExisteixTaula("[" & LastDatabase & "].dbo." & taulaServits) Then
        If Not ExisteixTaula("[" & LastDatabase & "_bak].dbo." & taulaServits) Then

            ExecutaComandaSql SqlCTS
            ExecutaComandaSql SqlTrace
            ExecutaComandaSql SqlTrigger
            
            taulaServits = "[" & LastDatabase & "].dbo." & taulaServits
        Else
            taulaServits = "[" & LastDatabase & "_bak].dbo." & taulaServits
        End If
    Else
        If Not ExisteixTaula("[" & LastDatabase & "].dbo." & TaulaServitsTrace) Then
            ExecutaComandaSql SqlTrace
        End If
        taulaServits = "[" & LastDatabase & "].dbo." & taulaServits
    End If

End Function

Function DonamNomTaulaNorma43() As String
   Dim sql As String
   
   DonamNomTaulaNorma43 = "Norma43"
   
   If Not ExisteixTaula("Norma43") Then
        sql = "CREATE TABLE [Norma43] ("
        sql = sql & "[IdFichero] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_Empresa] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_Banco] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_Oficina] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_numcuenta] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_Fi] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_Ff] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_SaldoInicial] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_NombreCliente] [nvarchar] (255)   NULL ,"
        sql = sql & "[Comu_CodigoCliente] [nvarchar] (255)   NULL ,"
        sql = sql & "[numlineaA] [numeric](18, 0) NULL ,"
        sql = sql & "[numlineaB] [numeric](18, 0) NULL ,"
        sql = sql & "[tipusRegistreA] [nvarchar] (2)   NULL ,"
        sql = sql & "[LliureA] [nvarchar] (4)   NULL ,"
        sql = sql & "[Oficina] [nvarchar] (4)   NULL ,"
        sql = sql & "[DataOperacio] [nvarchar] (6)   NULL ,"
        sql = sql & "[DataValor] [nvarchar] (6)   NULL ,"
        sql = sql & "[ConceptoComun] [nvarchar] (2)   NULL ,"
        sql = sql & "[ConceptoPropio] [nvarchar] (3)   NULL ,"
        sql = sql & "[DeveHaver] [nvarchar] (1)   NULL ,"
        sql = sql & "[Importe] [nvarchar] (14)   NULL ,"
        sql = sql & "[Documento] [nvarchar] (10)   NULL ,"
        sql = sql & "[Referencia] [nvarchar] (11)   NULL ,"
        sql = sql & "[Referencia2] [nvarchar] (16)   NULL ,"
        sql = sql & "[tipusRegistreB] [nvarchar] (2)   NULL ,"
        sql = sql & "[CodigoDato] [nvarchar] (2)   NULL ,"
        sql = sql & "[Concepto1] [nvarchar] (38)   NULL ,"
        sql = sql & "[Concepto2] [nvarchar] (38)   NULL ,"
        sql = sql & "[Concepto3] [nvarchar] (38)   NULL ,"
        sql = sql & "[Concepto4] [nvarchar] (38)   NULL ,"
        sql = sql & "[Concepto5] [nvarchar] (38)   NULL,"
        sql = sql & "[idNorma43] [nvarchar] (255),"
        sql = sql & "[hit_DataOperacio] [datetime] NULL,"
        sql = sql & "[hit_DataValor] [datetime] NULL,"
        sql = sql & "[hit_Importe] [float] NULL,"
        sql = sql & "[hit_Tipologia] [nvarchar] (255) NULL,"
        sql = sql & "[hit_Estado] [bit] NOT NULL CONSTRAINT [DF_Norma43_hit_Estado] DEFAULT (0) "
        
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
   End If

End Function



Function DonamNomTaulaNorma43Conta() As String
   Dim sql As String
   
   DonamNomTaulaNorma43Conta = "Norma43Conta"
   
   If Not ExisteixTaula(DonamNomTaulaNorma43Conta) Then
        sql = "CREATE TABLE [" & DonamNomTaulaNorma43Conta & "] ("
        sql = sql & "[idConta] [nvarchar](255) NULL, "
        sql = sql & "[idNorma43] [nvarchar](255) NULL, "
        sql = sql & "[valor] [nvarchar](255) NULL, "
        sql = sql & "[concepto] [nvarchar](255) NULL, "
        sql = sql & "[orden] [nvarchar](255) NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
   End If

End Function




Function SousNominaImportats() As String
   Dim sql As String
   SousNominaImportats = "SousNominaImportats"
   
   If Not ExisteixTaula("SousNominaImportats") Then
        sql = "create table " & SousNominaImportats & " ("
        sql = sql & "[IdFichero] [nvarchar](255) NULL, "
        sql = sql & "[Origen] [nvarchar](255) NULL, "
        sql = sql & "[CodiEmpresa] [nvarchar](255) NULL, "
        sql = sql & "[data] [nvarchar](255) NULL, "
        sql = sql & "[treb] [nvarchar](255) NULL, "
        sql = sql & "[i_liquid] [float] NULL, "
        sql = sql & "[I_Irpf] [float] NULL, "
        sql = sql & "[I_Tc1] [float] NULL, "
        sql = sql & "[I_Brut] [float] NULL, "
        sql = sql & "[i_SsEmp] [float] NULL, "
        sql = sql & "[i_SsTre] [float] NULL "
        sql = sql & ") ON [PRIMARY] "
        ExecutaComandaSql sql
   End If
   
   If Not ExisteixCamp("SousNominaImportats", "RetibEspecie") Then ExecutaComandaSql "ALTER TABLE SousNominaImportats ADD RetibEspecie [float] NULL "
   If Not ExisteixCamp("SousNominaImportats", "IrpfEspecies") Then ExecutaComandaSql "ALTER TABLE SousNominaImportats ADD IrpfEspecies [float] NULL "

End Function

Function SousNominaImportatsTmp() As String
   Dim sql As String
   SousNominaImportatsTmp = "SousNominaImportatsTmp"
   
   If Not ExisteixTaula("SousNominaImportatsTmp") Then
        sql = "create table " & SousNominaImportatsTmp & " ("
        sql = sql & "[IdFichero] [nvarchar](255) NULL, "
        sql = sql & "[Origen] [nvarchar](255) NULL, "
        sql = sql & "[CodiEmpresa] [nvarchar](255) NULL, "
        sql = sql & "[data] [nvarchar](255) NULL, "
        sql = sql & "[treb] [nvarchar](255) NULL, "
        sql = sql & "[i_liquid] [float] NULL, "
        sql = sql & "[I_Irpf] [float] NULL, "
        sql = sql & "[I_Tc1] [float] NULL, "
        sql = sql & "[I_Brut] [float] NULL, "
        sql = sql & "[i_SsEmp] [float] NULL, "
        sql = sql & "[i_SsTre] [float] NULL, "
        sql = sql & "[RetibEspecie] [float] NULL, "
        sql = sql & "[IrpfEspecies] [float] NULL "
        sql = sql & ") ON [PRIMARY] "
        ExecutaComandaSql sql
   End If

End Function






















Function NomTaulaAlbarans(data As Date) As String

   NomTaulaAlbarans = "V_Albarans_" & Format(data, "yyyy-mm")
   
End Function


Function NomTaulaFacturaData(data As Date) As String
   
   NomTaulaFacturaData = "Facturacio_" & Format(data, "yyyy-mm") & "_Data"
   
End Function
Function NomTaulaComptavilitzat(data As Date) As String
    Dim sql As String
   
   NomTaulaComptavilitzat = "Comptabilitzats_" & Year(data) & "-" & Right("00" & Month(data), 2)
   If Not ExisteixTaula(NomTaulaComptavilitzat) Then
        sql = "CREATE TABLE [" & NomTaulaComptavilitzat & "] ("
        sql = sql & " [id] [nvarchar] (255) NULL Default (NEWID()),"
        sql = sql & " [tipus] [nvarchar] (255) NULL,"
        sql = sql & " [numero] [decimal](18, 0) NULL,"
        sql = sql & " [id_doc] [nvarchar] (255) NULL,"
        sql = sql & " [data_doc] [datetime] NULL,"
        sql = sql & " [comptabilitzat] [nvarchar] (255) NULL,"
        sql = sql & " [data_comptab] [datetime] NULL,"
        sql = sql & " ) ON [PRIMARY]"
        ExecutaComandaSql sql
    End If
   
   
End Function

Function NomTaulaSecreAvisos() As String
    Dim sql As String
    
    NomTaulaSecreAvisos = "SecreAvisos"
    If Not ExisteixTaula(NomTaulaSecreAvisos) Then
        sql = "CREATE TABLE " & NomTaulaSecreAvisos & "  ( "
        sql = sql & " [Id]          [nvarchar] (255) NULL CONSTRAINT [DF_SecreAvisos_Id] DEFAULT (newid()),"
        sql = sql & " [Tipus]       [nvarchar] (255) NULL ,"
        sql = sql & " [TimeStamp]   [datetime]  NULL CONSTRAINT [DF_SecreAvisos_Ts] DEFAULT (GetDate()),"
        sql = sql & " [Usuari]      [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure1]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure2]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure3]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure4]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure5]     [nvarchar] (255) NULL )"
        ExecutaComandaSql sql
    End If
   
   
End Function


Function NomTaulaSecreAvisosTxt() As String
    Dim sql As String
    
    NomTaulaSecreAvisosTxt = "SecreAvisosTxt"
    If Not ExisteixTaula(NomTaulaSecreAvisosTxt) Then
        sql = "CREATE TABLE " & NomTaulaSecreAvisosTxt & "  ( "
        sql = sql & " [Id]          [nvarchar] (255) NULL CONSTRAINT [DF_SecreAvisosTxt_Id] DEFAULT (newid()),"
        sql = sql & " [Avisat]      [nvarchar] (255) NULL ,"
        sql = sql & " [Tipus]       [nvarchar] (255) NULL ,"
        sql = sql & " [TimeStamp]   [datetime]  NULL CONSTRAINT [DF_SecreAvisosTxt_Ts] DEFAULT (GetDate()),"
        sql = sql & " [Usuari]      [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure1]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure2]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure3]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure4]     [nvarchar] (255) NULL ,"
        sql = sql & " [Lliure5]     [nvarchar] (255) NULL )"
        
        ExecutaComandaSql sql
    End If
   
End Function

Function tablaFacturaProforma(D As Date) As String
    Dim sql As String, sqlFP As String

    tablaFacturaProforma = "ccFacturas_" & Year(D) & "_Iva"
    If Not ExisteixTaula(tablaFacturaProforma) Then
        sqlFP = "CREATE TABLE " & tablaFacturaProforma & " ("
        sqlFP = sqlFP & "   IdFactura nvarchar (255) NULL,"
        sqlFP = sqlFP & "   NumFactura float NULL,"
        sqlFP = sqlFP & "   EmpresaCodi nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Serie nvarchar (255) NULL,"
        sqlFP = sqlFP & "   DataInici datetime NULL,"
        sqlFP = sqlFP & "   DataFi datetime NULL,"
        sqlFP = sqlFP & "   DataFactura datetime NULL,"
        sqlFP = sqlFP & "   DataEmissio datetime NULL,"
        sqlFP = sqlFP & "   DataVenciment datetime NULL,"
        sqlFP = sqlFP & "   FormaPagament nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Total float NULL,"
        sqlFP = sqlFP & "   ClientCodi nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCodiFac nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientNom nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientNif nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientAdresa nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCp nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Tel nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Fax nvarchar (255) NULL,"
        sqlFP = sqlFP & "   eMail nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientLliure nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCiutat nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpNom nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpNif nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpAdresa nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpCp nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpTel nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpFax nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpeMail nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpLliure nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpCiutat nvarchar (255) NULL,"
        sqlFP = sqlFP & "   CampMercantil nvarchar (255) NULL,"
        sqlFP = sqlFP & "   BaseIva1 float NULL,"
        sqlFP = sqlFP & "   Iva1 float NULL,"
        sqlFP = sqlFP & "   BaseIva2 float NULL,"
        sqlFP = sqlFP & "   Iva2 float NULL,"
        sqlFP = sqlFP & "   BaseIva3 float NULL,"
        sqlFP = sqlFP & "   Iva3 float NULL,"
        sqlFP = sqlFP & "   BaseIva4 float NULL,"
        sqlFP = sqlFP & "   Iva4 float NULL,"
        sqlFP = sqlFP & "   BaseRec1 float NULL,"
        sqlFP = sqlFP & "   Rec1 float NULL,"
        sqlFP = sqlFP & "   BaseRec2 float NULL,"
        sqlFP = sqlFP & "   Rec2 float NULL,"
        sqlFP = sqlFP & "   BaseRec3 float NULL,"
        sqlFP = sqlFP & "   Rec3 float NULL,"
        sqlFP = sqlFP & "   BaseRec4 float NULL,"
        sqlFP = sqlFP & "   Rec4 float NULL,"
        sqlFP = sqlFP & "   valorIva1 float NULL,"
        sqlFP = sqlFP & "   valorIva2 float NULL,"
        sqlFP = sqlFP & "   valorIva3 float NULL,"
        sqlFP = sqlFP & "   valorIva4 float NULL,"
        sqlFP = sqlFP & "   valorRec1 float NULL,"
        sqlFP = sqlFP & "   valorRec2 float NULL,"
        sqlFP = sqlFP & "   valorRec3 float NULL,"
        sqlFP = sqlFP & "   valorRec4 float NULL,"
        sqlFP = sqlFP & "   IvaRec1 float NULL,"
        sqlFP = sqlFP & "   IvaRec2 float NULL,"
        sqlFP = sqlFP & "   IvaRec3 float NULL,"
        sqlFP = sqlFP & "   IvaRec4 float NULL,"
        sqlFP = sqlFP & "   Reservat nvarchar (255) NULL"
        sqlFP = sqlFP & "   )"
        ExecutaComandaSql sqlFP
    End If
   
End Function

Function tablaFacturaProformaAuto(D As Date) As String
    Dim sql As String, sqlFP As String

    tablaFacturaProformaAuto = "ccFacturas_" & Year(D) & "_Auto"
    If Not ExisteixTaula(tablaFacturaProformaAuto) Then
        sqlFP = "CREATE TABLE " & tablaFacturaProformaAuto & " ("
        sqlFP = sqlFP & "   IdFactura nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Data datetime NULL,"
        sqlFP = sqlFP & "   Usuari nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Comentari nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Autorizada bit NULL"
        sqlFP = sqlFP & "   )"
        
        ExecutaComandaSql sqlFP
    End If
   
End Function


Function tablaFacturaProformaE(D As Date, empresa As String) As String
    Dim sql As String, sqlFP As String

    tablaFacturaProformaE = empresa & ".dbo.ccFacturas_" & Year(D) & "_Iva"
    If Not ExisteixTaula(tablaFacturaProformaE) Then
        sqlFP = "CREATE TABLE " & tablaFacturaProformaE & " ("
        sqlFP = sqlFP & "   IdFactura nvarchar (255) NULL,"
        sqlFP = sqlFP & "   NumFactura nvarchar(255) NULL,"
        sqlFP = sqlFP & "   EmpresaCodi nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Serie nvarchar (255) NULL,"
        sqlFP = sqlFP & "   DataInici datetime NULL,"
        sqlFP = sqlFP & "   DataFi datetime NULL,"
        sqlFP = sqlFP & "   DataFactura datetime NULL,"
        sqlFP = sqlFP & "   DataEmissio datetime NULL,"
        sqlFP = sqlFP & "   DataVenciment datetime NULL,"
        sqlFP = sqlFP & "   FormaPagament nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Total float NULL,"
        sqlFP = sqlFP & "   ClientCodi nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCodiFac nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientNom nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientNif nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientAdresa nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCp nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Tel nvarchar (255) NULL,"
        sqlFP = sqlFP & "   Fax nvarchar (255) NULL,"
        sqlFP = sqlFP & "   eMail nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientLliure nvarchar (255) NULL,"
        sqlFP = sqlFP & "   ClientCiutat nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpNom nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpNif nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpAdresa nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpCp nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpTel nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpFax nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpeMail nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpLliure nvarchar (255) NULL,"
        sqlFP = sqlFP & "   EmpCiutat nvarchar (255) NULL,"
        sqlFP = sqlFP & "   CampMercantil nvarchar (255) NULL,"
        sqlFP = sqlFP & "   BaseIva1 float NULL,"
        sqlFP = sqlFP & "   Iva1 float NULL,"
        sqlFP = sqlFP & "   BaseIva2 float NULL,"
        sqlFP = sqlFP & "   Iva2 float NULL,"
        sqlFP = sqlFP & "   BaseIva3 float NULL,"
        sqlFP = sqlFP & "   Iva3 float NULL,"
        sqlFP = sqlFP & "   BaseIva4 float NULL,"
        sqlFP = sqlFP & "   Iva4 float NULL,"
        sqlFP = sqlFP & "   BaseRec1 float NULL,"
        sqlFP = sqlFP & "   Rec1 float NULL,"
        sqlFP = sqlFP & "   BaseRec2 float NULL,"
        sqlFP = sqlFP & "   Rec2 float NULL,"
        sqlFP = sqlFP & "   BaseRec3 float NULL,"
        sqlFP = sqlFP & "   Rec3 float NULL,"
        sqlFP = sqlFP & "   BaseRec4 float NULL,"
        sqlFP = sqlFP & "   Rec4 float NULL,"
        sqlFP = sqlFP & "   valorIva1 float NULL,"
        sqlFP = sqlFP & "   valorIva2 float NULL,"
        sqlFP = sqlFP & "   valorIva3 float NULL,"
        sqlFP = sqlFP & "   valorIva4 float NULL,"
        sqlFP = sqlFP & "   valorRec1 float NULL,"
        sqlFP = sqlFP & "   valorRec2 float NULL,"
        sqlFP = sqlFP & "   valorRec3 float NULL,"
        sqlFP = sqlFP & "   valorRec4 float NULL,"
        sqlFP = sqlFP & "   IvaRec1 float NULL,"
        sqlFP = sqlFP & "   IvaRec2 float NULL,"
        sqlFP = sqlFP & "   IvaRec3 float NULL,"
        sqlFP = sqlFP & "   IvaRec4 float NULL,"
        sqlFP = sqlFP & "   Reservat nvarchar (255) NULL"
        sqlFP = sqlFP & "   )"
        ExecutaComandaSql sqlFP
    End If
   
End Function



Function tablaFacturaProformaData(D As Date) As String
    Dim sql As String, sqlFP As String

    tablaFacturaProformaData = "ccFacturas_" & Year(D) & "_Data"
    If Not ExisteixTaula(tablaFacturaProformaData) Then
        sqlFP = "CREATE TABLE " & tablaFacturaProformaData & " ( "
        sqlFP = sqlFP & " IdFactura nvarchar (255) NULL, "
        sqlFP = sqlFP & " Data datetime NULL, "
        sqlFP = sqlFP & " Client nvarchar (255) NULL, "
        sqlFP = sqlFP & " Producte nvarchar (255) NULL, "
        sqlFP = sqlFP & " ProducteNom nvarchar (255) NULL, "
        sqlFP = sqlFP & " Acabat float NULL, "
        sqlFP = sqlFP & " Preu float NULL, "
        sqlFP = sqlFP & " Import float NULL, "
        sqlFP = sqlFP & " Desconte float NULL, "
        sqlFP = sqlFP & " TipusIva float NULL, "
        sqlFP = sqlFP & " Iva float NULL, "
        sqlFP = sqlFP & " rec float NULL, "
        sqlFP = sqlFP & " Referencia nvarchar (255) NULL, "
        sqlFP = sqlFP & " Servit float NULL, "
        sqlFP = sqlFP & " Tornat float NULL "
        sqlFP = sqlFP & "   )"
        ExecutaComandaSql sqlFP
    End If
End Function



Function tablaFacturaProformaDataE(D As Date, empresa As String) As String
    Dim sql As String, sqlFP As String

    tablaFacturaProformaDataE = empresa & ".dbo.ccFacturas_" & Year(D) & "_Data"
    If Not ExisteixTaula(tablaFacturaProformaDataE) Then
        sqlFP = "CREATE TABLE " & tablaFacturaProformaDataE & " ( "
        sqlFP = sqlFP & " IdFactura nvarchar (255) NULL, "
        sqlFP = sqlFP & " Data datetime NULL, "
        sqlFP = sqlFP & " Client nvarchar (255) NULL, "
        sqlFP = sqlFP & " Producte nvarchar (255) NULL, "
        sqlFP = sqlFP & " ProducteNom nvarchar (255) NULL, "
        sqlFP = sqlFP & " Acabat float NULL, "
        sqlFP = sqlFP & " Preu float NULL, "
        sqlFP = sqlFP & " Import float NULL, "
        sqlFP = sqlFP & " Desconte float NULL, "
        sqlFP = sqlFP & " TipusIva float NULL, "
        sqlFP = sqlFP & " Iva float NULL, "
        sqlFP = sqlFP & " rec float NULL, "
        sqlFP = sqlFP & " Referencia nvarchar (255) NULL, "
        sqlFP = sqlFP & " Servit float NULL, "
        sqlFP = sqlFP & " Tornat float NULL "
        sqlFP = sqlFP & "   )"
        ExecutaComandaSql sqlFP
    End If
End Function


Function NomTaulaRebutData(data As Date) As String

   NomTaulaRebutData = "Facturacio_" & Format(data, "yyyy-mm") & "_Reb"
   
End Function

Function NomTaulaFacturaIva(data As Date) As String

   NomTaulaFacturaIva = "Facturacio_" & Format(data, "yyyy-mm") & "_Iva"

End Function
Function NomTaulaFacturaReb(data As Date) As String

   NomTaulaFacturaReb = "Facturacio_" & Format(data, "yyyy-mm") & "_Reb"

End Function

Function NomTaulaRebuts(data As Date) As String

   NomTaulaRebuts = "Facturacio_" & Format(data, "yyyy-mm") & "_Reb"
   
End Function


Function NomTaulaAnulats(data As Date) As String

   NomTaulaAnulats = "V_Anulats_" & Format(data, "yyyy-mm")
   
End Function


Function NomTaulaLog(data As Date) As String

   NomTaulaLog = "V_Log_" & Format(data, "yyyy-mm")
   
End Function

Function NomTaulaDevol(data As Date) As String
   
   Dim sql As String
    
   NomTaulaDevol = "V_Tornat_" & Format(data, "yyyy-mm")
   
   If Not ExisteixTaula("[" & NomTaulaDevol & "]") Then
   
        sql = "CREATE TABLE [" & NomTaulaDevol & "] ("
        sql = sql & " [Botiga] [float] NULL, "
        sql = sql & " [Data] [datetime] NULL,"
        sql = sql & " [Dependenta] [float] NULL,"
        sql = sql & " [Num_tick] [float] NULL,"
        sql = sql & " [Estat] [nvarchar](25) NULL,"
        sql = sql & " [Plu] [float] NULL,"
        sql = sql & " [Quantitat] [float] NULL,"
        sql = sql & " [Import] [float] NULL,"
        sql = sql & " [Tipus_venta] [nvarchar](25) NULL)"

        ExecutaComandaSql sql
   End If
   
End Function

Function NomTaulaHoraris(data As Date) As String

   NomTaulaHoraris = "V_Horaris_" & Format(data, "yyyy-mm")
   If Not ExisteixTaula(NomTaulaHoraris) Then
      ExecutaComandaSql "Create Table [" & NomTaulaHoraris & "] (Botiga float,Data datetime,Dependenta float,Operacio [nvarchar] (25)) "
   End If
   
   
End Function

Function TaulaTicksRecapitulats(data As Date) As String
   Dim sql As String

   TaulaTicksRecapitulats = "Tiquets_Recapitulativa_" & Format(data, "yyyy")
   If Not ExisteixTaula(TaulaTicksRecapitulats) Then
        sql = "CREATE TABLE [dbo].[" & TaulaTicksRecapitulats & "] ("
        sql = sql & "[DataTick] [datetime] NOT NULL, "
        sql = sql & "[Botiga] [int] NOT NULL, "
        sql = sql & "[NumTick] [float] NOT NULL, "
        sql = sql & "[IdFactura] [nvarchar](255) NOT NULL, "
        sql = sql & "[DataFactura] [datetime] NOT NULL, "
        sql = sql & "[Serie] [nvarchar](15) NOT NULL, "
        sql = sql & "[NumFactura] [float] NOT NULL "
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
   End If
      
End Function
Function NomTaulaencarregs(data As Date) As String

   NomTaulaencarregs = "V_Encarre_" & Format(data, "yyyy-mm")
   
End Function
Function NomTaulaInventari(data As Date) As String
    Dim sql As String
    
   NomTaulaInventari = "V_Inventaris_" & Format(data, "yyyy-mm")
   
   If Not ExisteixTaula("[" & NomTaulaInventari & "]") Then
      sql = "Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25),Plu float,Quantitat float,Import float,Tipus_venta [nvarchar] (25) "
      ExecutaComandaSql "Create Table [" & NomTaulaInventari & "] (" & sql & ") "
   End If
   
   
End Function
Sub Obre_Base(Codi_Botiga As Double, data As Date)
'  Dim Sql As String
'
'  If Not ExisteixTaula(NomTaulaVentas(Data)) Then CreaTaulesDadesTpv Data
'
'
'  Set PreuProducte = Db.CreateQuery("", "Select Preu From Articles Where CodiGenetic  = ? ")
'
'  Set Registre_Horaris = Db.CreateQuery("", "Insert Into [" & NomTaulaHoraris(Data) & "] (Botiga,Data,Dependenta,Operacio) Values (?,?,?,?)")
'  Set Registre_Tickets = Db.CreateQuery("", "Insert Into [" & NomTaulaVentas(Data) & "] (Botiga,Data,Dependenta,Num_tick,Plu,Quantitat,Import,Tipus_venta,Estat) Values (?,?,?,?,?,?,?,?,?)")
'  Set Registre_Movimen = Db.CreateQuery("", "Insert Into [" & NomTaulaMovi(Data) & "] (Botiga,Data,Dependenta,Tipus_moviment,Import,Motiu) Values (?,?,?,?,?,?)")
'  Set Registre_Pa_Sec = Db.CreateQuery("", "Insert Into [" & NomTaulaDevol(Data) & "] (botiga,Data,Dependenta,Num_tick,Plu,Quantitat,Import,Tipus_venta,Estat) Values (?,?,?,?,?,?,?,?,?)")
'
'  Sql = "Insert Into [" & NomTaulaAnulats(Data) & "] Select * From [" & NomTaulaVentas(Data) & "] "
'  Sql = Sql & "WHERE [Num_tick] = ? AND "
'  Sql = Sql & "DATA >= ? AND "
'  Sql = Sql & "DATA <= ? "
'  Set Esb1 = Db.CreateQuery("", Sql)
'
'  Sql = "Delete [" & NomTaulaVentas(Data) & "] "
'  Sql = Sql & "WHERE [Num_tick] = ? AND "
'  Sql = Sql & "DATA >= ? AND "
'  Sql = Sql & "DATA <= ? "
'  Set Esb2 = Db.CreateQuery("", Sql)
'
'  Db.BeginTrans
'
'  ExecutaComandaSql "Delete [" & NomTaulaVentas(Data) & "] Where Botiga = " & Codi_Botiga & " And Day(Data) = " & Day(Data) & " "
'  ExecutaComandaSql "Delete [" & NomTaulaAnulats(Data) & "] Where Botiga = " & Codi_Botiga & " And Day(Data) = " & Day(Data) & " "
'  ExecutaComandaSql "Delete [" & NomTaulaDevol(Data) & "] Where Botiga = " & Codi_Botiga & " And Day(Data) = " & Day(Data) & " "
'  ExecutaComandaSql "Delete [" & NomTaulaMovi(Data) & "] Where Botiga = " & Codi_Botiga & " And Day(Data) = " & Day(Data) & " "
'  ExecutaComandaSql "Delete [" & NomTaulaHoraris(Data) & "] Where Botiga = " & Codi_Botiga & " And Day(Data) = " & Day(Data) & " "

End Sub

Sub afegeig_registre_ticket(Hora, dependenta, tick_num, Plu, Quantitat, import, Tipus_venta)
   Dim Data_i_Hora
   
   If existeix_ultim_ticket Then
      Registre_Tickets.rdoParameters(0) = u_t_botiga_en_conversio
      Registre_Tickets.rdoParameters(1) = u_t_data_i_hora
      Registre_Tickets.rdoParameters(2) = u_t_dependenta
      Registre_Tickets.rdoParameters(3) = u_t_tick_num
      Registre_Tickets.rdoParameters(4) = u_t_plu
      Registre_Tickets.rdoParameters(5) = u_t_quantitat
      Registre_Tickets.rdoParameters(6) = u_t_import
      Registre_Tickets.rdoParameters(7) = u_t_Tipus_venta
      Registre_Tickets.rdoParameters(8) = u_t_estat
      Registre_Tickets.Execute
      
      existeix_ultim_ticket = False
   End If
   
   Data_i_Hora = DateValue(data_en_conversio) + TimeValue(Hora)
   existeix_ultim_ticket = True
   u_t_botiga_en_conversio = botiga_en_conversio
   u_t_data_i_hora = Data_i_Hora
   u_t_dependenta = dependenta
   u_t_tick_num = tick_num
   u_t_plu = Plu
   u_t_quantitat = Quantitat
   u_t_import = import
   u_t_Tipus_venta = Tipus_venta
   u_t_estat = "No impres"
   u_t_data_i_hora_UltimHuma = u_t_data_i_hora
  
End Sub
Sub afegeig_registre_devolucions(Hora, dependenta, Plu, Quantitat, import, Tipus_venta)
   Dim Data_i_Hora
   
   Registre_Pa_Sec.rdoParameters(0) = u_t_botiga_en_conversio
   Registre_Pa_Sec.rdoParameters(1) = DateValue(data_en_conversio) + TimeValue(Hora)
   Registre_Pa_Sec.rdoParameters(2) = dependenta
   Registre_Pa_Sec.rdoParameters(3) = 0
   Registre_Pa_Sec.rdoParameters(4) = Plu
   Registre_Pa_Sec.rdoParameters(5) = Quantitat
   Registre_Pa_Sec.rdoParameters(6) = import
   Registre_Pa_Sec.rdoParameters(7) = "S"
   Registre_Pa_Sec.rdoParameters(8) = " "
   Registre_Pa_Sec.Execute
  
End Sub

Sub afegeig_registre_ultim_ticket()
   
   If existeix_ultim_ticket Then
      Registre_Tickets.rdoParameters(0) = u_t_botiga_en_conversio
      Registre_Tickets.rdoParameters(1) = u_t_data_i_hora
      Registre_Tickets.rdoParameters(2) = u_t_dependenta
      Registre_Tickets.rdoParameters(3) = u_t_tick_num
      Registre_Tickets.rdoParameters(4) = u_t_plu
      Registre_Tickets.rdoParameters(5) = u_t_quantitat
      Registre_Tickets.rdoParameters(6) = u_t_import
      Registre_Tickets.rdoParameters(7) = u_t_Tipus_venta
      Registre_Tickets.rdoParameters(8) = u_t_estat
      Registre_Tickets.Execute

      existeix_ultim_ticket = False
   End If
   
End Sub

Sub afegeig_registre_horari(Hora, dependenta, Operacio)
   
   Registre_Horaris.rdoParameters(0) = botiga_en_conversio
   Registre_Horaris.rdoParameters(1) = DateValue(data_en_conversio) + TimeValue(Hora)
   Registre_Horaris.rdoParameters(2) = dependenta
   Registre_Horaris.rdoParameters(3) = Operacio
   Registre_Horaris.Execute

End Sub
Sub afegeig_registre_moviments(Hora, dependenta, Operacio, import, Motiu)
   Dim Data_i_Hora, Clau

   Clau = ""
   If Operacio = "W" Then Clau = "Canvi"
   If Operacio = "O" Then Clau = "Motius Po"
   If Operacio = "A" Then Clau = "Motius Ra"
   
   Registre_Movimen.rdoParameters(0) = botiga_en_conversio
   Registre_Movimen.rdoParameters(1) = DateValue(data_en_conversio) + TimeValue(Hora)
   Registre_Movimen.rdoParameters(2) = dependenta
   Registre_Movimen.rdoParameters(3) = Operacio
   Registre_Movimen.rdoParameters(4) = import
   Registre_Movimen.rdoParameters(5) = TradueixMotiu(Motiu, Clau)
   Registre_Movimen.Execute

End Sub


Function Converteix(Fil As String) As Integer
   Dim cadena As String, cade As String
   Dim Tamany, Segu As Integer
   Dim pos_punt  As Integer
   Dim Portem As Double, Total As Double, Bones As Double, Dolentes() As String
   Dim CaixaOberta As Boolean
   
   dos_opened = False
   Handel1 = FreeFile
   Open DIRECTORI_TEMPORAL & "\" & Fil For Input As Handel1
   My_DoEvents
   Tamany = LOF(Handel1)
   botiga_en_conversio = File_CodiBotiga(NomFileCurt(Fil))
   data_en_conversio = file_data(NomFileCurt(Fil))
   Ultim_ticket = ""
   CaixaOberta = False
'On Error GoTo 0

   If Tamany > 0 Then
      Obre_Base botiga_en_conversio, data_en_conversio
      My_DoEvents
      anterior_hora = 0
      FilActual = "Interpretant " & Format(data_en_conversio, "dddd dd-mm-yyyy") & "(" & botiga_en_conversio & ")"
      InformaConv 0
      
      Line Input #Handel1, cade
      Portem = Len(cade)
      InitVectorDependentes
      elTemps_NoCorre_Endarrera = 0

      ReDim HoraRealZeta(0)
      While Not EOF(Handel1)
         Line Input #Handel1, cadena
         Portem = Portem + Len(cadena)
         If Estat = E_cal_plegar Then
            Segu = MsgBox("Abortar l' operació implica RECOMENSAR el proces , ABORTEM ?", 4 + 32, "Atenció")
            If Segu = 6 Then GoTo user_diu_fi2
            Estat = E_treballant
         End If
On Error GoTo fitcher_dolent
         Select Case Left(cadena, 1)
            Case MARCA_VENTA_AL_COMPTAT
                 tracta_venta_al_comptat cadena
                 CaixaOberta = True
            Case MARCA_BON_DIA
                 tracta_bon_dia cadena
            Case MARCA_ENTRADA_DEPENDENTA
                 tracta_entrada_dependenta cadena
                 CaixaOberta = True
            Case MARCA_CAMVI
                 tracta_camvi cadena
                 CaixaOberta = True
            Case MARCA_PLEGA_DEP
                 tracta_plega_dep cadena
                 CaixaOberta = True
            Case MARCA_MOVIMENT_PO
                 tracta_moviment_po cadena
                 CaixaOberta = True
            Case MARCA_MOVIMENT_RA
                 tracta_moviment_ra cadena
                 CaixaOberta = True
            Case MARCA_POSAR_HORA
                 tracta_posar_hora cadena
            Case MARCA_POSAR_DATA
                 tracta_posar_data cadena
            Case MARCA_DESCONTE1
                 tracta_desconte1 cadena
                 CaixaOberta = True
            Case MARCA_DESCONTE2
                 tracta_desconte2 cadena
                 CaixaOberta = True
            Case MARCA_CAIXA_TANCADA
                 afegeig_registre_ultim_ticket
                 tracta_caixa_tancada cadena
                 CaixaOberta = False
            Case MARCA_IMPRESIO_TIQUET
                 tracta_impresio_ticket cadena
            Case MARCA_ERROR_AL_MARCAR
                 tracta_error_al_nmarcar cadena
                 CaixaOberta = True
            Case MARCA_PA_SEC
                 tracta_pa_sec cadena
            Case MARCA_DEP_SURT_UN_MOMENT
                 tracta_dep_surt_un_moment cadena
                 CaixaOberta = True
            Case MARCA_TORNA_DEPENDENTA
                 tracta_torna_dependenta cadena
                 CaixaOberta = True
            Case MARCA_FACTURAT
                 tracta_facturat cadena
                 CaixaOberta = True
         End Select
     Wend
     afegeig_registre_ultim_ticket
     esborra_errors DIRECTORI_TEMPORAL & Fil
     
     Db.CommitTrans
     
     If CaixaOberta Then tracta_caixa_no_tancada
     
     ActualizaPanSeco data_en_conversio, botiga_en_conversio
     
   End If
   
   Close Handel1
   If dos_opened Then Close Handel2
   dos_opened = False

   Converteix = 1
   Exit Function
fitcher_dolent:
   Db.RollbackTrans
   tanca_base
   Close Handel1
   If dos_opened Then Close Handel2
   dos_opened = False
   Converteix = 2
   Exit Function
user_diu_fi2:
   Db.RollbackTrans
   tanca_base
   Close Handel1
   If dos_opened Then Close Handel2
   dos_opened = False
   Converteix = 3
   Exit Function
End Function

Function digital(Minuts)
   digital = Int(Minuts / 60) & ":" & (Minuts - (Int(Minuts / 60) * 60))
End Function



Function digital2(Minuts)
    
   digital2 = Format(Minuts / 60, "####0.00") 'Int(Minuts / 60) & ".." & Format((Minuts - (Int(Minuts / 60) * 60)), "00")
   
   If Minuts = 0 Then digital2 = ""
   
End Function


Function digital3(Minuts)
    
   digital3 = Format(Minuts / 60, "####0.00") 'Int(Minuts / 60) & ".." & Format((Minuts - (Int(Minuts / 60) * 60)), "00")
   
   If Minuts = 0 Then digital3 = "0"
   
End Function




Sub InformaConv(Hora)
   Dim H As Integer, m As Integer, K As Double

'   k = 2 * (Hora / (24 * 60)) - 1
'   frmMain.Sol.Move (frmMain.BaseDia.Width) * ((Sin(k) + 1) / 2), frmMain.BaseDia.Height + frmMain.Sol.Height - (frmMain.BaseDia.Height * ((Cos(k) * 2) - 1))

   H = Int(Hora / 60)
   m = Hora Mod 60
   
   If Not ControlEstat Is Nothing Then ControlEstat.Caption = FilActual & "  " & Format(H, "00") & ":" & Format(m, "00")
   My_DoEvents
'   frmMain.Hora = Format(H, "00") & ":" & Format(M, "00")
   My_DoEvents

End Sub
Function parser(Cad, pos)
   Dim old_pos
   old_pos = pos + 1
   pos = InStr(old_pos, Cad, "+")
   If pos = 0 Then
      parser = ""
      pos = Len(Cad)
   Else
      parser = Mid(Cad, old_pos, pos - old_pos)
   End If
End Function

Function PosanHi(L As String, i As Integer) As String
   PosanHi = Left((L & Space(i)), i)
End Function

Function Tabula(L As String, i As Integer) As String
   Tabula = Right((Space(i) & L), i)
End Function

Sub tanca_base()
End Sub


Sub TancaFile(f)
   On Error Resume Next
   Close f
   
End Sub

Sub tracta_bon_dia(cadena)
    Dim pos As Integer, dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    
    afegeig_registre_horari Hora, dependenta, OP_BON_DIA
    
End Sub
Sub tracta_caixa_no_tancada()
   Dim i As Integer, Hora As Date, Zeta As Double, dia As String, Myset As rdoResultset, MyQuery As rdoQuery, MySql As String
   Dim data_inici As Date, Data_Fi As Date, clients As Double, dependenta As Double
   
   Hora = CVDate(Day(data_en_conversio) & "-" & Month(data_en_conversio) & "-" & Year(data_en_conversio) & " 01:01:00 ")
   dependenta = 0

On Error Resume Next
   Set Myset = Db.OpenResultset("Select Max(Data) As D FROM [" & NomTaulaHoraris(data_en_conversio) & "] Where Day(Data) = " & Day(data_en_conversio) & " ")
   If Not Myset.EOF > 0 Then If Hora < Myset(0) Then Hora = Myset(0)
   Myset.Close
   
   Set Myset = Db.OpenResultset("Select Max(Data) As D FROM [" & NomTaulaVentas(data_en_conversio) & "] Where Day(Data) = " & Day(data_en_conversio) & " ")
   If Not Myset.EOF > 0 Then If Hora < Myset(0) Then Hora = Myset(0)
   Myset.Close
   
   Set Myset = Db.OpenResultset("Select Max(Data) As D FROM [" & NomTaulaMovi(data_en_conversio) & "] Where Day(Data) = " & Day(data_en_conversio) & " ")
   If Not Myset.EOF > 0 Then If Hora < Myset(0) Then Hora = Myset(0)
   Myset.Close
On Error GoTo 0

   Hora = DateAdd("n", 1, Hora)
   
   For i = 1 To NumDependentesALaBotiga
      afegeig_registre_horari Hora, DependentesALaBotiga(i), OP_PLEGA_DEPENDENTA
   Next
   
   dia = Format(DateValue(data_en_conversio), "dd-mm-yy")
   Set Myset = Db.OpenResultset("SELECT Data FROM [" & NomTaulaMovi(data_en_conversio) & "] WHERE [Tipus_Moviment] = 'Z' Order By Data")
   
   i = 0
   data_inici = CVDate(Day(dia) & "-" & Month(dia) & "-" & Year(dia) & " 01:01:00")
   Data_Fi = Hora
   
   While Not Myset.EOF
      data_inici = Myset("Data")
      Myset.MoveNext
   Wend
   Myset.Close
   
   data_inici = DateAdd("n", 1, data_inici)
   
   MySql = "SELECT Sum(Import) AS Sumat,Count(Import) AS Contat FROM [" & NomTaulaVentas(data_en_conversio) & "] "
   MySql = MySql & "WHERE Data >= ? AND "
   MySql = MySql & " Data <= ? "
   
   Set MyQuery = Db.CreateQuery("", MySql)
   MyQuery.rdoParameters(0) = data_inici
   MyQuery.rdoParameters(1) = Data_Fi
   
   Set Myset = MyQuery.OpenResultset
   
   Zeta = 0
   clients = 0
   If Not Myset.EOF Then
      If Not IsNull(Myset("Sumat")) Then Zeta = Myset("Sumat")
      If Not IsNull(Myset("Contat")) Then clients = Myset("Contat")
   End If
   Myset.Close
   
   MySql = "SELECT Dependenta FROM [" & NomTaulaHoraris(data_en_conversio) & "] "
   MySql = MySql & "WHERE Data >= ? AND "
   MySql = MySql & "Data <= ?       AND "
   MySql = MySql & "Operacio = 'E' "
   MySql = MySql & "Order By Data "
   
   Set MyQuery = Db.CreateQuery("", MySql)
   
   MyQuery.rdoParameters(0) = data_inici
   MyQuery.rdoParameters(1) = Data_Fi
   
   Set Myset = MyQuery.OpenResultset
   
   If Not Myset.EOF Then
      dependenta = Myset("Dependenta")
   End If
   Myset.Close
   
   Hora = DateAdd("n", 1, Hora)
   
   afegeig_registre_moviments Hora, dependenta, OP_DESCUADRE, 0, " "
   afegeig_registre_moviments Hora, dependenta, OP_NOMBRE_CLIENTS, clients, " "
   afegeig_registre_moviments Hora, dependenta, OP_CAIXA_TANCADA, Zeta, " "
   
   
End Sub

Sub tracta_caixa_tancada(cadena)
Dim i As Integer
'   Z945+46+1+364+5+765+25+3600+100+21100+500+7000+1000+11000+2000+12000+5000+20000+10000+0+
'   Z883+46+1+264+5+1750+25+8025+100+24800+500+0+1000+34000+2000+0+5000+5000+10000+0+
    
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Total As Double, Pts As Double
    Dim moneda As String
    Dim Hora As String
    Dim son_les As Integer
    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    
    ReDim Preserve HoraRealZeta(UBound(HoraRealZeta) + 1)
    HoraRealZeta(UBound(HoraRealZeta)) = son_les
    
'    InformaConv son_les
    Hora = digital(son_les)
    
    dependenta = Val(parser(cadena, pos))
    moneda = parser(cadena, pos)
    
    Total = 0
    Do
      Pts = Val(parser(cadena, pos))
      Total = Total + Pts
      afegeig_registre_moviments Hora, dependenta, OP_ENTRADA_CANVI, Pts, moneda
      moneda = parser(cadena, pos)
    Loop While Len(moneda)
    
'    InformaConv son_les
    
    For i = 1 To NumDependentesALaBotiga
       afegeig_registre_horari Hora, DependentesALaBotiga(i), OP_PLEGA_DEPENDENTA
    Next
    
    InitVectorDependentes
    
End Sub


Sub tracta_camvi(cadena)
    Dim pos As Integer, dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
'    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    
    afegeig_registre_moviments Hora, dependenta, OP_CAMVI, 0, " "
   
End Sub

Sub tracta_dep_surt_un_moment(cadena)
    Dim pos As Integer, dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))

    afegeig_registre_horari Hora, dependenta, OP_DEP_SURT_UN_MOMENT

End Sub

Sub tracta_desconte1(cadena)
    Dim pos As Integer
    Dim Plu As Double, tick_num As Double, Quantitat As Double, a_pes As Double
    Dim dependenta As Integer, productes As Integer, i As Integer
    Dim Hora As String
    Dim quants As Double, import As Double
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    productes = Val(parser(cadena, pos))
    tick_num = Val(parser(cadena, pos))
    

    
    For i = 1 To productes
      Quantitat = Val(parser(cadena, pos))
      a_pes = Val(parser(cadena, pos))
      Plu = parser(cadena, pos)
      quants = Quantitat / a_pes
      import = quants * Article_Plu_Preu(Plu)

      afegeig_registre_ticket Hora, dependenta, tick_num, Plu, quants, import - Int(import * 0.1), OP_DESCONTE1
    Next
    
End Sub

Sub tracta_desconte2(cadena)
    Dim pos As Integer
    Dim Plu As Double, tick_num As Double, Quantitat As Double, a_pes As Double
    Dim dependenta As Integer, productes As Integer, i As Integer
    Dim Hora As String
    Dim quants As Double, import As Double
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    productes = parser(cadena, pos)
    tick_num = Val(parser(cadena, pos))

    
    For i = 1 To productes
      Quantitat = Val(parser(cadena, pos))
      a_pes = Val(parser(cadena, pos))
      Plu = parser(cadena, pos)
      
      quants = Quantitat / a_pes
      import = quants * Article_Plu_Preu(Plu)
      
      afegeig_registre_ticket Hora, dependenta, tick_num, Plu, quants, import - Int(import * 0.3), OP_DESCONTE2
    Next
    
End Sub

Sub tracta_entrada_dependenta(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))

    afegeig_registre_horari Hora, dependenta, OP_ENTRADA_DEPENDENTA
    NumDependentesALaBotiga = NumDependentesALaBotiga + 1
    ReDim Preserve DependentesALaBotiga(NumDependentesALaBotiga)
    DependentesALaBotiga(NumDependentesALaBotiga) = dependenta
End Sub

Sub tracta_error_al_nmarcar(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
'    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    
'    afegeig_registre OP_ERROR_AL_MARCAR, Hora, dependenta No es don en aquest fitcher
    
End Sub

Sub tracta_errors_ticket(cadena As String)
    
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer
    Dim tick_num As Double
    Dim MySql As String

    pos = 1
    son_les = Val(parser(cadena, pos))
    u_t_data_i_hora_UltimHuma = son_les
    dependenta = Val(parser(cadena, pos))
    tick_num = Val(parser(cadena, pos))
    
    num_tickets_esborrats = num_tickets_esborrats + 1
    ReDim Preserve tickets_esborrats(num_tickets_esborrats)
    tickets_esborrats(num_tickets_esborrats) = tick_num
    
End Sub

Sub tracta_errors_z(cadena, NumZeta As Integer)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Date
    Dim tick_num As Double
    Dim MySql As String
    Dim i As Integer
    Dim son_les2 As Integer
    Dim Caixa, descuadre, clients

    If Not ControlEstat Is Nothing Then ControlEstat.Caption = FilActual & "  Esborrant Tickets Erronis "
    My_DoEvents
    pos = 1
    son_les2 = Val(parser(cadena, pos))
    
'    InformaConv son_les2

    Pen_Ultima_Zeta = Ultima_Zeta
    Ultima_Zeta = DateValue(data_en_conversio) + TimeValue(digital(son_les2))

    For i = 1 To num_tickets_esborrats
       If Not ControlEstat Is Nothing Then ControlEstat.Caption = FilActual & " Esborrant Erronis " & i
       My_DoEvents
       
       Esb1.rdoParameters(0) = tickets_esborrats(i)
       Esb1.rdoParameters(1) = Pen_Ultima_Zeta
       Esb1.rdoParameters(2) = Ultima_Zeta
       Esb1.Execute
       
       Esb2.rdoParameters(0) = tickets_esborrats(i)
       Esb2.rdoParameters(1) = Pen_Ultima_Zeta
       Esb2.rdoParameters(2) = Ultima_Zeta
       Esb2.Execute
    
    Next
    num_tickets_esborrats = 0
    ReDim tickets_esborrats(num_tickets_esborrats)
    If NumZeta <= UBound(HoraRealZeta) Then
       If son_les2 < HoraRealZeta(NumZeta) Then son_les2 = HoraRealZeta(NumZeta)
    End If
    NumZeta = NumZeta + 1
    
    Hora = digital(son_les2)
    
    dependenta = Val(parser(cadena, pos))
    Caixa = parser(cadena, pos)
    descuadre = parser(cadena, pos)
    clients = parser(cadena, pos)
    
    afegeig_registre_moviments Hora, dependenta, OP_DESCUADRE, descuadre, " "
    afegeig_registre_moviments Hora, dependenta, OP_NOMBRE_CLIENTS, clients, " "
    afegeig_registre_moviments Hora, dependenta, OP_CAIXA_TANCADA, Caixa, " "
        
End Sub

Sub tracta_facturat(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    
'    afegeig_registre OP_FACTURAT, Hora, dependenta, 0, 0, 0, 0
    
End Sub

Sub tracta_impresio_ticket(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    pos = 1
    'hora = digital(Val(parser(cadena, pos)))
    dependenta = Val(parser(cadena, pos))
    marca_ultim_ticket_com_a_impres
End Sub

Sub tracta_moviment_po(cadena)
    Dim pos As Integer
    Dim Motiu As Double, import As Double
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    Motiu = Val(parser(cadena, pos))
    import = Val(parser(cadena, pos))
    
    afegeig_registre_moviments Hora, dependenta, OP_MOVIMENT_PO, import, Motiu
    
End Sub

Sub tracta_moviment_ra(cadena)
    Dim pos As Integer
    Dim Motiu As Double, import As Double
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

' O863+80+0+56000+

    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    Motiu = Val(parser(cadena, pos))
    import = Val(parser(cadena, pos))
    
    afegeig_registre_moviments Hora, dependenta, OP_MOVIMENT_RA, import, Motiu
    
End Sub

Sub tracta_pa_sec(cadena)
    Dim pos As Integer
    Dim Plu As Double, tick_num As Double, Quantitat As Double, a_pes As Double
    Dim dependenta As Integer, productes As Integer, i As Integer
    Dim Hora As String
    Dim quants As Double, import As Double
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    productes = Val(parser(cadena, pos))
    tick_num = Val(parser(cadena, pos))
    
    For i = 1 To productes
      Quantitat = Val(parser(cadena, pos))
      a_pes = Val(parser(cadena, pos))
      Plu = parser(cadena, pos)
      quants = Quantitat / a_pes
          
      import = Int(quants * Article_Plu_Preu(Plu))
            
      afegeig_registre_devolucions Hora, dependenta, Plu, quants, import, OP_PA_SEC
      
    Next
    
    
End Sub

Sub tracta_plega_dep(cadena)
    Dim pos As Integer
    Dim dependenta As Integer, Motiu As Double, import As Double
    Dim Hora As String
    Dim son_les As Integer
Dim i As Integer
    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    
    afegeig_registre_horari Hora, dependenta, OP_PLEGA_DEPENDENTA
    i = 1
    While i <= NumDependentesALaBotiga
       If DependentesALaBotiga(i) = dependenta Then
          While i < NumDependentesALaBotiga
             DependentesALaBotiga(i) = DependentesALaBotiga(i + 1)
             i = i + 1
          Wend
          NumDependentesALaBotiga = NumDependentesALaBotiga - 1
          ReDim Preserve DependentesALaBotiga(NumDependentesALaBotiga)
          i = NumDependentesALaBotiga + 1
       End If
       i = i + 1
    Wend
End Sub

Sub tracta_posar_data(cadena)
    Dim pos As Integer
    Dim dependenta As Integer, Motiu As Double, import As Double
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))

    afegeig_registre_horari Hora, dependenta, OP_POSAR_DATA
End Sub

Sub tracta_posar_hora(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))

    afegeig_registre_horari Hora, dependenta, OP_POSAR_HORA
End Sub

Sub tracta_torna_dependenta(cadena)
    Dim pos As Integer
    Dim dependenta As Integer
    Dim Hora As String
    Dim son_les As Integer
    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    afegeig_registre_horari Hora, dependenta, OP_TORNA_DEPENDENTA
End Sub

Sub tracta_venta_al_comptat(cadena)
    Dim pos As Integer
    Dim Plu As Double, tick_num As Double, Quantitat As Double, a_pes As Double
    Dim dependenta As Integer, productes As Integer, i As Integer
    Dim Hora As String
    Dim quants As Double, import As Double
    Dim son_les As Integer

    
    pos = 1
    son_les = Val(parser(cadena, pos))
    If son_les < elTemps_NoCorre_Endarrera Then
       son_les = elTemps_NoCorre_Endarrera
    End If
    elTemps_NoCorre_Endarrera = son_les
    u_t_data_i_hora_UltimHuma = son_les
    InformaConv son_les
    Hora = digital(son_les)
    dependenta = Val(parser(cadena, pos))
    productes = Val(parser(cadena, pos))
    tick_num = Val(parser(cadena, pos))
    
    For i = 1 To productes
      Quantitat = Val(parser(cadena, pos))
      a_pes = Val(parser(cadena, pos))
      Plu = parser(cadena, pos)
      quants = Quantitat / a_pes
          
      import = Int(quants * Article_Plu_Preu(Plu))
            
      afegeig_registre_ticket Hora, dependenta, tick_num, Plu, quants, import, OP_VENTA_AL_COMPTAT
    Next
    
End Sub


Function Article_Plu_Preu(Plu As Double) As Double
   
   Article_Plu_Preu = 0
   If Plu > 0 And Plu < UBound(HashPreus) Then Article_Plu_Preu = HashPreus(Plu)
   
End Function

Function TradueixMotiu(Motiu, Clau) As String
   
   TradueixMotiu = "Motiu : " & Motiu
   
'   If CLau = "" Then Exit Function
'
'   If CLau = "Canvi" Then
'      TradueixMotiu = "En : " & Motiu
'      Exit Function
'   End If
'
'   TaulaMotius.Seek "=", Motiu
'   While Not TaulaMotius.EOF
'      If Not TaulaMotius("Ordre") = Motiu Then Exit Function
'      If TaulaMotius("Clau") = CLau Then
'         TradueixMotiu = TaulaMotius("Concepte")
'         Exit Function
'      End If
'      TaulaMotius.MoveNext
'   Wend
   
End Function


