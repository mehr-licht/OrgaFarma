Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
Imports System.Data.OleDb
Imports System.IO
Imports System.Text

Public Class Form1
    Inherits Form
    Dim aceitarduplicados As Boolean = True
    Dim showcode As Boolean
    Dim quantosDCIa1nosP As Single
    Dim quantosDCIa2nosP As Single
    Dim quantosDCIa3nosP As Single
    Dim quantosDCIa4nosP As Single
    Dim quantosDCIa1 As Single
    Dim quantosDCIa2 As Single
    Dim quantosDCIa3 As Single
    Dim quantosDCIa4 As Single
    Dim quantosDCIp1 As Single
    Dim quantosDCIp2 As Single
    Dim quantosDCIp3 As Single
    Dim quantosDCIp4 As Single
    Dim inicializado As Boolean
    Dim comparado As Boolean
    Dim mal As New Boolean
    Dim avisoL As New Boolean
    Dim avisoH As New Boolean
    Dim marcawarning As New Boolean
    Dim codigoantes As New Integer
    Dim inicialReader As StreamReader
    Dim inicialWriter As StreamWriter
    Dim FileReader As StreamReader
    Dim FileWriter As StreamWriter
    Dim monthReader As StreamReader
    Dim monthWriter As StreamWriter
    Dim lastorder = 25
    Dim resultado1 As New embalagem.cruzamento
    Dim resultado2 As New embalagem.cruzamento
    Dim resultado3 As New embalagem.cruzamento
    Dim resultado4 As New embalagem.cruzamento
    Dim cruzamento11 As New embalagem.cruzamento
    Dim cruzamento12 As New embalagem.cruzamento
    Dim cruzamento13 As New embalagem.cruzamento
    Dim cruzamento14 As New embalagem.cruzamento
    Dim cruzamento21 As New embalagem.cruzamento
    Dim cruzamento22 As New embalagem.cruzamento
    Dim cruzamento23 As New embalagem.cruzamento
    Dim cruzamento24 As New embalagem.cruzamento
    Dim cruzamento31 As New embalagem.cruzamento
    Dim cruzamento32 As New embalagem.cruzamento
    Dim cruzamento33 As New embalagem.cruzamento
    Dim cruzamento34 As New embalagem.cruzamento
    Dim cruzamento41 As New embalagem.cruzamento
    Dim cruzamento42 As New embalagem.cruzamento
    Dim cruzamento43 As New embalagem.cruzamento
    Dim cruzamento44 As New embalagem.cruzamento
    Dim troca1 As String = "0"
    Dim troca2 As String = "0"
    Dim troca3 As String = "0"
    Dim troca4 As String = "0"
    Dim oa1 As New embalagem.ordem
    Dim oa2 As New embalagem.ordem
    Dim oa3 As New embalagem.ordem
    Dim oa4 As New embalagem.ordem
    Dim oad12 As New embalagem.ordem
    Dim oad13 As New embalagem.ordem
    Dim oad14 As New embalagem.ordem
    Dim oad23 As New embalagem.ordem
    Dim oad24 As New embalagem.ordem
    Dim oad34 As New embalagem.ordem
    Dim oat123 As New embalagem.ordem
    Dim oat134 As New embalagem.ordem
    Dim opt134 As New embalagem.ordem
    Dim oat124 As New embalagem.ordem
    Dim oat234 As New embalagem.ordem
    Dim oaq As New embalagem.ordem
    Dim op1 As New embalagem.ordem
    Dim op2 As New embalagem.ordem
    Dim op3 As New embalagem.ordem
    Dim op4 As New embalagem.ordem
    Dim opd12 As New embalagem.ordem
    Dim opd13 As New embalagem.ordem
    Dim opd14 As New embalagem.ordem
    Dim opd23 As New embalagem.ordem
    Dim opd24 As New embalagem.ordem
    Dim opd34 As New embalagem.ordem
    Dim opq As New embalagem.ordem
    Dim opt123 As New embalagem.ordem
    Dim opt124 As New embalagem.ordem
    Dim opt234 As New embalagem.ordem
    Dim resultadomostrado As Boolean
    Dim vshe As Single
    Dim vshe1 As Single
    Dim vshe2 As Single
    Dim vshe3 As Single
    Dim vshe4 As Single
    Dim filtrosoisolados As Boolean = True
    Dim filtromarcamarcadci As Boolean = True
    'filtromarcamarcadci falso desactiva a nova verificação de troca de marca quando não há genérico e só foi prescrito 1 medicamento (no p1a1234)
    Dim filtrolab As Boolean = False
    'filtrolab falso desactiva os antigos erros de troca de lab, marca2gen, gen2marca, marca-marca
    Dim agrupado As Boolean = False
    Dim ghtext As String
    Dim ghtext1 As String
    Dim unidose As Boolean = False
    Dim tectocomp As Double = tectocomp
    Dim tectocomp2 As Double = tectocomp
    Dim manipcomp As Double = 0.3
    Dim manipcomp2 As Double = 0.3
    Dim mudarentidadelimpa As Boolean = True
    Dim haport As Boolean = False
    Dim port1474 As Boolean = False
    'Dim farmacia As Integer
    Dim tempcalc As Double
    Dim mostrarsubvia As Boolean = True
    Dim mostrarmarca As Boolean = True
    Dim mostrarcoddif As Boolean = False
    Dim mostrartrocalab As Boolean = True
    Dim mostrarqtyinferior As Boolean = False
    Dim taxaQuant As Double = 60 / 56
    Dim taxapr As Short = 1
    Dim conjunto As Short
    Dim amarelo As Boolean
    Dim vermelho As Boolean
    Dim verde As Boolean
    Dim varcruza1 As Short
    Dim varcruza2 As Short
    Dim varcruza3 As Short
    Dim varcruza4 As Short
    Dim varcruzp1 As Short
    Dim varcruzp2 As Short
    Dim varcruzp3 As Short
    Dim varcruzp4 As Short
    Dim varlabelcruz1 As String
    Dim varlabelcruz2 As String
    Dim varlabelcruz3 As String
    Dim varlabelcruz4 As String
    Dim portcomp As Double
    Dim portcomp1 As Double
    Dim portcomp2 As Double
    Dim portcomp3 As Double
    Dim portcomp4 As Double
    Dim portcompdois As Double
    Dim portcomp11 As Double
    Dim portcomp22 As Double
    Dim portcomp33 As Double
    Dim portcomp44 As Double
    Dim intermedio As Double
    Dim intermedio2 As Double
    Dim pvp1v As String
    Dim pvp2v As String
    Dim pvp3v As String
    Dim pvp4v As String
    Dim pvp1val As Double
    Dim pvp2val As Double
    Dim pvp3val As Double
    Dim pvp4val As Double
    Dim comp1v As String
    Dim comp2v As String
    Dim comp3v As String
    Dim comp4v As String
    Dim comp1val As Double
    Dim comp2val As Double
    Dim comp3val As Double
    Dim comp4val As Double
    Dim pvp11v As String
    Dim pvp22v As String
    Dim pvp33v As String
    Dim pvp44v As String
    Dim pvp11val As Double
    Dim pvp22val As Double
    Dim pvp33val As Double
    Dim pvp44val As Double
    Dim comp11v As String
    Dim comp22v As String
    Dim comp33v As String
    Dim comp44v As String
    Dim comp11val As Double
    Dim comp22val As Double
    Dim comp33val As Double
    Dim comp44val As Double
    Dim pr As Double
    Dim pr1 As Double
    Dim pr2 As Double
    Dim pr3 As Double
    Dim pr4 As Double
    Dim pr11 As Double
    Dim pr22 As Double
    Dim pr33 As Double
    Dim pr44 As Double
    Dim organismo As String
    Dim gen As Boolean
    Dim comp As Double
    Dim compdois As Double
    Dim portimedio As String
    Dim mostrado91 As Boolean
    Dim portec4a1 As String
    Dim portec4b1 As String
    Dim genec41 As String
    Dim mostracnpem As String
    Dim mostracnpem1 As String
    Dim mostranome1 As String
    Dim mostradci1 As String
    Dim mostraforma1 As String
    Dim mostradose1 As String
    Dim mostraqty1 As String
    Dim mostragh1 As String
    Dim mostracomp1 As String
    Dim mostracompgen1 As String
    Dim mostraports1 As String
    Dim mostrado9 As Boolean
    Dim portec4a As String
    Dim portec4b As String
    Dim genec4 As String
    Dim mostranome As String
    Dim mostradci As String
    Dim mostraforma As String
    Dim mostradose As String
    Dim mostraqty As String
    Dim mostragh As String
    Dim mostracomp As String
    Dim mostracompgen As String
    Dim mostraports As String
    Dim codigorow As basededadosDataSet.infarmedRow
    Dim codigoarray As New ArrayList
    Dim codigo41 As New meds
    Dim codigo4 As New meds
    Dim novado1 As Boolean
    Dim novado As Boolean

    Dim linha1(0 To 8)
    Dim linha2(0 To 8)
    Dim linha3(0 To 8)
    Dim linha4(0 To 8)

    Dim p1, p2, p3, p4, a1, a2, a3, a4 As Integer

    Dim descomp1mostrado As Boolean
    Dim descomp2mostrado As Boolean
    Dim descomp3mostrado As Boolean
    Dim descomp4mostrado As Boolean

    Dim Prescrito1 As New meds
    Dim Prescrito2 As New meds
    Dim Prescrito3 As New meds
    Dim Prescrito4 As New meds

    Dim Aviado1 As New meds
    Dim Aviado2 As New meds
    Dim Aviado3 As New meds
    Dim Aviado4 As New meds

    Dim vazio As New meds

    'usado na avaliação keypress para saber em que controlo está o foco
    Dim foco As String

    'as variaveis A e P contam quantas embalagens foram (A)viadas e quantas foram (P)rescritas
    Dim A As Short
    Dim P As Short
    Dim AA As Short
    Dim PP As Short

    Dim procurado1 As Boolean
    Dim procurado2 As Boolean
    Dim procurado3 As Boolean
    Dim procurado4 As Boolean

    Dim infarmedTA As New basededadosDataSetTableAdapters.infarmedTableAdapter

    Dim DS As New basededadosDataSet

    Dim row As DataRow

    'esta variavel guarda o número de linhas lido da base de dados
    Dim infarmedTAs As Integer = infarmedTA.Fill(DS.infarmed)

    Dim p1row As basededadosDataSet.infarmedRow
    Dim p2row As basededadosDataSet.infarmedRow
    Dim p3row As basededadosDataSet.infarmedRow
    Dim p4row As basededadosDataSet.infarmedRow
    Dim a1row As basededadosDataSet.infarmedRow
    Dim a2row As basededadosDataSet.infarmedRow
    Dim a3row As basededadosDataSet.infarmedRow
    Dim a4row As basededadosDataSet.infarmedRow

    Dim a1array As New ArrayList
    Dim a2array As New ArrayList
    Dim a3array As New ArrayList
    Dim a4array As New ArrayList

    Dim p1array As New ArrayList
    Dim p2array As New ArrayList
    Dim p3array As New ArrayList
    Dim p4array As New ArrayList

    Dim a1p1 As New avaliacao
    Dim a2p1 As New avaliacao
    Dim a3p1 As New avaliacao
    Dim a4p1 As New avaliacao
    Dim a1p2 As New avaliacao
    Dim a2p2 As New avaliacao
    Dim a3p2 As New avaliacao
    Dim a4p2 As New avaliacao
    Dim a1p3 As New avaliacao
    Dim a2p3 As New avaliacao
    Dim a3p3 As New avaliacao
    Dim a4p3 As New avaliacao
    Dim a1p4 As New avaliacao
    Dim a2p4 As New avaliacao
    Dim a3p4 As New avaliacao
    Dim a4p4 As New avaliacao
    Dim av1 As New avaliacao
    Dim av2 As New avaliacao
    Dim av3 As New avaliacao
    Dim av4 As New avaliacao

    Dim grupoP1 As Short = 0
    Dim grupoP2 As Short = 0
    Dim grupoP3 As Short = 0
    Dim grupoP4 As Short = 0
    Dim grupoA1 As Short = 0
    Dim grupoA2 As Short = 0
    Dim grupoA3 As Short = 0
    Dim grupoA4 As Short = 0
    Dim grupoP1dci As String
    Dim grupoP2dci As String
    Dim grupoP3dci As String
    Dim grupoP4dci As String
    Dim grupoA1dci As String
    Dim grupoA2dci As String
    Dim grupoA3dci As String
    Dim grupoA4dci As String

    Dim a1_4250 As Boolean
    Dim a2_4250 As Boolean
    Dim a3_4250 As Boolean
    Dim a4_4250 As Boolean
    Dim a1_1234 As Boolean
    Dim a2_1234 As Boolean
    Dim a3_1234 As Boolean
    Dim a4_1234 As Boolean
    Dim a1_14123 As Boolean
    Dim a2_14123 As Boolean
    Dim a3_14123 As Boolean
    Dim a4_14123 As Boolean
    Dim a1_1474ad As Boolean
    Dim a2_1474ad As Boolean
    Dim a3_1474ad As Boolean
    Dim a4_1474ad As Boolean
    Dim a1_1474nl As Boolean
    Dim a2_1474nl As Boolean
    Dim a3_1474nl As Boolean
    Dim a4_1474nl As Boolean
    Dim a1_21094 As Boolean
    Dim a2_21094 As Boolean
    Dim a3_21094 As Boolean
    Dim a4_21094 As Boolean
    Dim a1_10279 As Boolean
    Dim a2_10279 As Boolean
    Dim a3_10279 As Boolean
    Dim a4_10279 As Boolean
    Dim a1_10910 As Boolean
    Dim a2_10910 As Boolean
    Dim a3_10910 As Boolean
    Dim a4_10910 As Boolean
    Dim limpo As Boolean
    Dim cordefundo As Color

    'Prepara para ler nova receita - limpa tudo (valores a zero, labels em branco e sem fundo) e foco na primeira caixa
    Sub inicializar()
a1:
a2:     On Error GoTo MOSTRARERRO
a3:     If inicializado = False Then
            resultadomostrado = False
            read_only()
            resultadototal.Font = New Font(resultadototal.Font, FontStyle.Bold)
            resultadototal.SelectionAlignment = HorizontalAlignment.Center
            PictureBox1.BringToFront()
            resultadototal.Hide()
            PictureBox1.Show()
            cordefundo = SystemColors.GradientInactiveCaption
            PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
a4:         comparado = False
a5:         mal = False
            avisoL = False
            avisoH = False
            marcawarning = False
a6:         codigoantes = 0
a7:         resultadototal.Text = ""
            justificacao.Text = ""
a8:         op1.nDCI = vbEmpty
a9:         op1.pvpmenos2 = vbEmpty
a10:        op1.pvptop5 = vbEmpty
a11:        op1.pvpmenos1top5 = vbEmpty
            op1.pvpmenos2top5 = vbEmpty
            op1.pvpun = vbEmpty
            op1.pvpmenos1un = vbEmpty
            op1.pvpmenos2un = vbEmpty
            op1.excepa = vbEmpty
            op1.excepb = vbEmpty
            op1.excepc = vbEmpty
            op1.prescrito = vbEmpty
            op1.cruzacom = vbEmpty
            op1.result = vbEmpty

            op2.nDCI = vbEmpty
            op2.pvpmenos2 = vbEmpty
            op2.pvptop5 = vbEmpty
            op2.pvpmenos1top5 = vbEmpty
            op2.pvpmenos2top5 = vbEmpty
            op2.pvpun = vbEmpty
            op2.pvpmenos1un = vbEmpty
            op2.pvpmenos2un = vbEmpty
            op2.excepa = vbEmpty
            op2.excepb = vbEmpty
            op2.excepc = vbEmpty
            op2.prescrito = vbEmpty
            op2.cruzacom = vbEmpty
            op2.result = vbEmpty


            op3.nDCI = vbEmpty
            op3.pvpmenos2 = vbEmpty
            op3.pvptop5 = vbEmpty
            op3.pvpmenos1top5 = vbEmpty
            op3.pvpmenos2top5 = vbEmpty
            op3.pvpun = vbEmpty
            op3.pvpmenos1un = vbEmpty
            op3.pvpmenos2un = vbEmpty
            op3.excepa = vbEmpty
            op3.excepb = vbEmpty
            op3.excepc = vbEmpty
            op3.prescrito = vbEmpty
            op3.cruzacom = vbEmpty
            op3.result = vbEmpty

            op4.nDCI = vbEmpty
            op4.pvpmenos2 = vbEmpty
            op4.pvptop5 = vbEmpty
            op4.pvpmenos1top5 = vbEmpty
            op4.pvpmenos2top5 = vbEmpty
            op4.pvpun = vbEmpty
            op4.pvpmenos1un = vbEmpty
            op4.pvpmenos2un = vbEmpty
            op4.excepa = vbEmpty
            op4.excepb = vbEmpty
            op4.excepc = vbEmpty
            op4.prescrito = vbEmpty
            op4.cruzacom = vbEmpty
            op4.result = vbEmpty

            oa1.code = vbEmpty
            oa1.dci = vbEmpty
            oa1.forma = vbEmpty
            oa1.dose = vbEmpty
            oa1.qty = vbEmpty
            oa1.comp = vbEmpty
            oa1.gh = vbEmpty
            oa1.gen = vbEmpty
            oa1.lab = vbEmpty
            oa1.pvp = vbEmpty
            oa1.pr = vbEmpty
            oa1.d4250 = vbEmpty
            oa1.d1234 = vbEmpty
            oa1.d10279 = vbEmpty
            oa1.d10280 = vbEmpty
            oa1.d21094 = vbEmpty
            oa1.lei6 = vbEmpty
            oa1.d14123 = vbEmpty
            oa1.dci_obr = vbEmpty
            oa1.trocamarca = vbEmpty
            oa1.top5 = vbEmpty
            oa1.pvpmenos1 = vbEmpty
            oa1.CNPEM = vbEmpty
            oa1.porCNPEM = vbEmpty
            oa1.duplicado = vbEmpty
            oa1.nDCI = vbEmpty
            oa1.pvpmenos2 = vbEmpty
            oa1.pvptop5 = vbEmpty
            oa1.pvpmenos1top5 = vbEmpty
            oa1.pvpmenos2top5 = vbEmpty
            oa1.pvpun = vbEmpty
            oa1.pvpmenos1un = vbEmpty
            oa1.pvpmenos2un = vbEmpty
            oa1.excepa = vbEmpty
            oa1.excepb = vbEmpty
            oa1.excepc = vbEmpty
            oa1.prescrito = vbEmpty
            oa1.cruzacom = vbEmpty
            oa1.result = vbEmpty

            oa2.code = vbEmpty
            oa2.dci = vbEmpty
            oa2.forma = vbEmpty
            oa2.dose = vbEmpty
            oa2.qty = vbEmpty
            oa2.comp = vbEmpty
            oa2.gh = vbEmpty
            oa2.gen = vbEmpty
            oa2.lab = vbEmpty
            oa2.pvp = vbEmpty
            oa2.pr = vbEmpty
            oa2.d4250 = vbEmpty
            oa2.d1234 = vbEmpty
            oa2.d10279 = vbEmpty
            oa2.d10280 = vbEmpty
            oa2.d21094 = vbEmpty
            oa2.lei6 = vbEmpty
            oa2.d14123 = vbEmpty
            oa2.dci_obr = vbEmpty
            oa2.trocamarca = vbEmpty
            oa2.top5 = vbEmpty
            oa2.pvpmenos1 = vbEmpty
            oa2.CNPEM = vbEmpty
            oa2.porCNPEM = vbEmpty
            oa2.duplicado = vbEmpty
            oa2.nDCI = vbEmpty
            oa2.pvpmenos2 = vbEmpty
            oa2.pvptop5 = vbEmpty
            oa2.pvpmenos1top5 = vbEmpty
            oa2.pvpmenos2top5 = vbEmpty
            oa2.pvpun = vbEmpty
            oa2.pvpmenos1un = vbEmpty
            oa2.pvpmenos2un = vbEmpty
            oa2.excepa = vbEmpty
            oa2.excepb = vbEmpty
            oa2.excepc = vbEmpty
            oa2.prescrito = vbEmpty
            oa2.cruzacom = vbEmpty
            oa2.result = vbEmpty

            oa3.code = vbEmpty
            oa3.dci = vbEmpty
            oa3.forma = vbEmpty
            oa3.dose = vbEmpty
            oa3.qty = vbEmpty
            oa3.comp = vbEmpty
            oa3.gh = vbEmpty
            oa3.gen = vbEmpty
            oa3.lab = vbEmpty
            oa3.pvp = vbEmpty
            oa3.pr = vbEmpty
            oa3.d4250 = vbEmpty
            oa3.d1234 = vbEmpty
            oa3.d10279 = vbEmpty
            oa3.d10280 = vbEmpty
            oa3.d21094 = vbEmpty
            oa3.lei6 = vbEmpty
            oa3.d14123 = vbEmpty
            oa3.dci_obr = vbEmpty
            oa3.trocamarca = vbEmpty
            oa3.top5 = vbEmpty
            oa3.pvpmenos1 = vbEmpty
            oa3.CNPEM = vbEmpty
            oa3.porCNPEM = vbEmpty
            oa3.duplicado = vbEmpty
            oa3.nDCI = vbEmpty
            oa3.pvpmenos2 = vbEmpty
            oa3.pvptop5 = vbEmpty
            oa3.pvpmenos1top5 = vbEmpty
            oa3.pvpmenos2top5 = vbEmpty
            oa3.pvpun = vbEmpty
            oa3.pvpmenos1un = vbEmpty
            oa3.pvpmenos2un = vbEmpty
            oa3.excepa = vbEmpty
            oa3.excepb = vbEmpty
            oa3.excepc = vbEmpty
            oa3.prescrito = vbEmpty
            oa3.cruzacom = vbEmpty
            oa3.result = vbEmpty

            oa4.code = vbEmpty
            oa4.dci = vbEmpty
            oa4.forma = vbEmpty
            oa4.dose = vbEmpty
            oa4.qty = vbEmpty
            oa4.comp = vbEmpty
            oa4.gh = vbEmpty
            oa4.gen = vbEmpty
            oa4.lab = vbEmpty
            oa4.pvp = vbEmpty
            oa4.pr = vbEmpty
            oa4.d4250 = vbEmpty
            oa4.d1234 = vbEmpty
            oa4.d10279 = vbEmpty
            oa4.d10280 = vbEmpty
            oa4.d21094 = vbEmpty
            oa4.lei6 = vbEmpty
            oa4.d14123 = vbEmpty
            oa4.dci_obr = vbEmpty
            oa4.trocamarca = vbEmpty
            oa4.top5 = vbEmpty
            oa4.pvpmenos1 = vbEmpty
            oa4.CNPEM = vbEmpty
            oa4.porCNPEM = vbEmpty
            oa4.duplicado = vbEmpty
            oa4.nDCI = vbEmpty
            oa4.pvpmenos2 = vbEmpty
            oa4.pvptop5 = vbEmpty
            oa4.pvpmenos1top5 = vbEmpty
            oa4.pvpmenos2top5 = vbEmpty
            oa4.pvpun = vbEmpty
            oa4.pvpmenos1un = vbEmpty
            oa4.pvpmenos2un = vbEmpty
            oa4.excepa = vbEmpty
            oa4.excepb = vbEmpty
            oa4.excepc = vbEmpty
            oa4.prescrito = vbEmpty
            oa4.cruzacom = vbEmpty
            oa4.result = vbEmpty

            cruzamento41.code = vbEmpty
            cruzamento41.dci = vbEmpty
            cruzamento41.nome = vbEmpty
            cruzamento41.forma = vbEmpty
            cruzamento41.dose = vbEmpty
            cruzamento41.qty = vbEmpty
            cruzamento41.comp = vbEmpty
            cruzamento41.gh = vbEmpty
            cruzamento41.pvp = vbEmpty
            cruzamento41.pr = vbEmpty
            cruzamento41.gen = vbEmpty
            cruzamento41.lab = vbEmpty
            cruzamento41.d4250 = vbEmpty
            cruzamento41.d1234 = vbEmpty
            cruzamento41.d21094 = vbEmpty
            cruzamento41.d10279 = vbEmpty
            cruzamento41.d10280 = vbEmpty
            cruzamento41.d10910 = vbEmpty
            cruzamento41.d14123 = vbEmpty
            cruzamento41.top5 = vbEmpty
            cruzamento41.dci_obr = vbEmpty
            cruzamento41.lei6 = vbEmpty
            cruzamento41.pvpmenos1 = vbEmpty
            cruzamento41.pvpmenos2 = vbEmpty
            cruzamento41.trocamarca = vbEmpty
            cruzamento41.CNPEM = vbEmpty
            cruzamento41.nDCI = vbEmpty  '?
            cruzamento41.excepa = vbEmpty 'se 
            cruzamento41.excepb = vbEmpty
            cruzamento41.excepc = vbEmpty
            cruzamento41.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento41.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento41.xgh = vbEmpty 'se existem ambos
            cruzamento41.xcnpem = vbEmpty 'se existem ambos
            cruzamento41.xnDCI = vbEmpty 'se existem ambos
            cruzamento41.porDCImesmoCNPEM = vbEmpty
            cruzamento41.porDCIdifCNPEM = vbEmpty
            cruzamento41.porMARCAmesmoCNPEM = vbEmpty
            cruzamento41.porMARCAdifCNPEM = vbEmpty
            cruzamento41.ranking = vbEmpty
            cruzamento41.anulado = vbEmpty
            cruzamento41.qualP = vbEmpty
            cruzamento41.marcadifsemhavergens = vbEmpty
            cruzamento42.code = vbEmpty
            cruzamento42.dci = vbEmpty
            cruzamento42.nome = vbEmpty
            cruzamento42.forma = vbEmpty
            cruzamento42.dose = vbEmpty
            cruzamento42.qty = vbEmpty
            cruzamento42.comp = vbEmpty
            cruzamento42.gh = vbEmpty
            cruzamento42.pvp = vbEmpty
            cruzamento42.pr = vbEmpty
            cruzamento42.gen = vbEmpty
            cruzamento42.lab = vbEmpty
            cruzamento42.d4250 = vbEmpty
            cruzamento42.d1234 = vbEmpty
            cruzamento42.d21094 = vbEmpty
            cruzamento42.d10279 = vbEmpty
            cruzamento42.d10280 = vbEmpty
            cruzamento42.d10910 = vbEmpty
            cruzamento42.d14123 = vbEmpty
            cruzamento42.top5 = vbEmpty
            cruzamento42.dci_obr = vbEmpty
            cruzamento42.lei6 = vbEmpty
            cruzamento42.pvpmenos1 = vbEmpty
            cruzamento42.pvpmenos2 = vbEmpty
            cruzamento42.trocamarca = vbEmpty
            cruzamento42.CNPEM = vbEmpty
            cruzamento42.nDCI = vbEmpty  '?
            cruzamento42.excepa = vbEmpty 'se 
            cruzamento42.excepb = vbEmpty
            cruzamento42.excepc = vbEmpty
            cruzamento42.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento42.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento42.xgh = vbEmpty 'se existem ambos
            cruzamento42.xcnpem = vbEmpty 'se existem ambos
            cruzamento42.xnDCI = vbEmpty 'se existem ambos
            cruzamento42.porDCImesmoCNPEM = vbEmpty
            cruzamento42.porDCIdifCNPEM = vbEmpty
            cruzamento42.porMARCAmesmoCNPEM = vbEmpty
            cruzamento42.porMARCAdifCNPEM = vbEmpty
            cruzamento42.ranking = vbEmpty
            cruzamento42.anulado = vbEmpty
            cruzamento42.qualP = vbEmpty
            cruzamento42.marcadifsemhavergens = vbEmpty

            cruzamento43.code = vbEmpty
            cruzamento43.dci = vbEmpty
            cruzamento43.nome = vbEmpty
            cruzamento43.forma = vbEmpty
            cruzamento43.dose = vbEmpty
            cruzamento43.qty = vbEmpty
            cruzamento43.comp = vbEmpty
            cruzamento43.gh = vbEmpty
            cruzamento43.pvp = vbEmpty
            cruzamento43.pr = vbEmpty
            cruzamento43.gen = vbEmpty
            cruzamento43.lab = vbEmpty
            cruzamento43.d4250 = vbEmpty
            cruzamento43.d1234 = vbEmpty
            cruzamento43.d21094 = vbEmpty
            cruzamento43.d10279 = vbEmpty
            cruzamento43.d10280 = vbEmpty
            cruzamento43.d10910 = vbEmpty
            cruzamento43.d14123 = vbEmpty
            cruzamento43.top5 = vbEmpty
            cruzamento43.dci_obr = vbEmpty
            cruzamento43.lei6 = vbEmpty
            cruzamento43.pvpmenos1 = vbEmpty
            cruzamento43.pvpmenos2 = vbEmpty
            cruzamento43.trocamarca = vbEmpty
            cruzamento43.CNPEM = vbEmpty
            cruzamento43.nDCI = vbEmpty  '?
            cruzamento43.excepa = vbEmpty 'se 
            cruzamento43.excepb = vbEmpty
            cruzamento43.excepc = vbEmpty
            cruzamento43.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento43.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento43.xgh = vbEmpty 'se existem ambos
            cruzamento43.xcnpem = vbEmpty 'se existem ambos
            cruzamento43.xnDCI = vbEmpty 'se existem ambos
            cruzamento43.porDCImesmoCNPEM = vbEmpty
            cruzamento43.porDCIdifCNPEM = vbEmpty
            cruzamento43.porMARCAmesmoCNPEM = vbEmpty
            cruzamento43.porMARCAdifCNPEM = vbEmpty
            cruzamento43.ranking = vbEmpty
            cruzamento43.anulado = vbEmpty
            cruzamento43.qualP = vbEmpty
            cruzamento43.marcadifsemhavergens = vbEmpty

            cruzamento44.code = vbEmpty
            cruzamento44.dci = vbEmpty
            cruzamento44.nome = vbEmpty
            cruzamento44.forma = vbEmpty
            cruzamento44.dose = vbEmpty
            cruzamento44.qty = vbEmpty
            cruzamento44.comp = vbEmpty
            cruzamento44.gh = vbEmpty
            cruzamento44.pvp = vbEmpty
            cruzamento44.pr = vbEmpty
            cruzamento44.gen = vbEmpty
            cruzamento44.lab = vbEmpty
            cruzamento44.d4250 = vbEmpty
            cruzamento44.d1234 = vbEmpty
            cruzamento44.d21094 = vbEmpty
            cruzamento44.d10279 = vbEmpty
            cruzamento44.d10280 = vbEmpty
            cruzamento44.d10910 = vbEmpty
            cruzamento44.d14123 = vbEmpty
            cruzamento44.top5 = vbEmpty
            cruzamento44.dci_obr = vbEmpty
            cruzamento44.lei6 = vbEmpty
            cruzamento44.pvpmenos1 = vbEmpty
            cruzamento44.pvpmenos2 = vbEmpty
            cruzamento44.trocamarca = vbEmpty
            cruzamento44.CNPEM = vbEmpty
            cruzamento44.nDCI = vbEmpty  '?
            cruzamento44.excepa = vbEmpty 'se 
            cruzamento44.excepb = vbEmpty
            cruzamento44.excepc = vbEmpty
            cruzamento44.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento44.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento44.xgh = vbEmpty 'se existem ambos
            cruzamento44.xcnpem = vbEmpty 'se existem ambos
            cruzamento44.xnDCI = vbEmpty 'se existem ambos
            cruzamento44.porDCImesmoCNPEM = vbEmpty
            cruzamento44.porDCIdifCNPEM = vbEmpty
            cruzamento44.porMARCAmesmoCNPEM = vbEmpty
            cruzamento44.porMARCAdifCNPEM = vbEmpty
            cruzamento44.ranking = vbEmpty
            cruzamento44.anulado = vbEmpty
            cruzamento44.qualP = vbEmpty
            cruzamento44.marcadifsemhavergens = vbEmpty


            cruzamento31.code = vbEmpty
            cruzamento31.dci = vbEmpty
            cruzamento31.nome = vbEmpty
            cruzamento31.forma = vbEmpty
            cruzamento31.dose = vbEmpty
            cruzamento31.qty = vbEmpty
            cruzamento31.comp = vbEmpty
            cruzamento31.gh = vbEmpty
            cruzamento31.pvp = vbEmpty
            cruzamento31.pr = vbEmpty
            cruzamento31.gen = vbEmpty
            cruzamento31.lab = vbEmpty
            cruzamento31.d4250 = vbEmpty
            cruzamento31.d1234 = vbEmpty
            cruzamento31.d21094 = vbEmpty
            cruzamento31.d10279 = vbEmpty
            cruzamento31.d10280 = vbEmpty
            cruzamento31.d10910 = vbEmpty
            cruzamento31.d14123 = vbEmpty
            cruzamento31.top5 = vbEmpty
            cruzamento31.dci_obr = vbEmpty
            cruzamento31.lei6 = vbEmpty
            cruzamento31.pvpmenos1 = vbEmpty
            cruzamento31.pvpmenos2 = vbEmpty
            cruzamento31.trocamarca = vbEmpty
            cruzamento31.CNPEM = vbEmpty
            cruzamento31.nDCI = vbEmpty  '?
            cruzamento31.excepa = vbEmpty 'se 
            cruzamento31.excepb = vbEmpty
            cruzamento31.excepc = vbEmpty
            cruzamento31.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento31.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento31.xgh = vbEmpty 'se existem ambos
            cruzamento31.xcnpem = vbEmpty 'se existem ambos
            cruzamento31.xnDCI = vbEmpty 'se existem ambos
            cruzamento31.porDCImesmoCNPEM = vbEmpty
            cruzamento31.porDCIdifCNPEM = vbEmpty
            cruzamento31.porMARCAmesmoCNPEM = vbEmpty
            cruzamento31.porMARCAdifCNPEM = vbEmpty
            cruzamento31.ranking = vbEmpty
            cruzamento31.anulado = vbEmpty
            cruzamento31.qualP = vbEmpty
            cruzamento31.marcadifsemhavergens = vbEmpty
            cruzamento32.code = vbEmpty
            cruzamento32.dci = vbEmpty
            cruzamento32.nome = vbEmpty
            cruzamento32.forma = vbEmpty
            cruzamento32.dose = vbEmpty
            cruzamento32.qty = vbEmpty
            cruzamento32.comp = vbEmpty
            cruzamento32.gh = vbEmpty
            cruzamento32.pvp = vbEmpty
            cruzamento32.pr = vbEmpty
            cruzamento32.gen = vbEmpty
            cruzamento32.lab = vbEmpty
            cruzamento32.d4250 = vbEmpty
            cruzamento32.d1234 = vbEmpty
            cruzamento32.d21094 = vbEmpty
            cruzamento32.d10279 = vbEmpty
            cruzamento32.d10280 = vbEmpty
            cruzamento32.d10910 = vbEmpty
            cruzamento32.d14123 = vbEmpty
            cruzamento32.top5 = vbEmpty
            cruzamento32.dci_obr = vbEmpty
            cruzamento32.lei6 = vbEmpty
            cruzamento32.pvpmenos1 = vbEmpty
            cruzamento32.pvpmenos2 = vbEmpty
            cruzamento32.trocamarca = vbEmpty
            cruzamento32.CNPEM = vbEmpty
            cruzamento32.nDCI = vbEmpty  '?
            cruzamento32.excepa = vbEmpty 'se 
            cruzamento32.excepb = vbEmpty
            cruzamento32.excepc = vbEmpty
            cruzamento32.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento32.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento32.xgh = vbEmpty 'se existem ambos
            cruzamento32.xcnpem = vbEmpty 'se existem ambos
            cruzamento32.xnDCI = vbEmpty 'se existem ambos
            cruzamento32.porDCImesmoCNPEM = vbEmpty
            cruzamento32.porDCIdifCNPEM = vbEmpty
            cruzamento32.porMARCAmesmoCNPEM = vbEmpty
            cruzamento32.porMARCAdifCNPEM = vbEmpty
            cruzamento32.ranking = vbEmpty
            cruzamento32.anulado = vbEmpty
            cruzamento32.qualP = vbEmpty
            cruzamento32.marcadifsemhavergens = vbEmpty

            cruzamento33.code = vbEmpty
            cruzamento33.dci = vbEmpty
            cruzamento33.nome = vbEmpty
            cruzamento33.forma = vbEmpty
            cruzamento33.dose = vbEmpty
            cruzamento33.qty = vbEmpty
            cruzamento33.comp = vbEmpty
            cruzamento33.gh = vbEmpty
            cruzamento33.pvp = vbEmpty
            cruzamento33.pr = vbEmpty
            cruzamento33.gen = vbEmpty
            cruzamento33.lab = vbEmpty
            cruzamento33.d4250 = vbEmpty
            cruzamento33.d1234 = vbEmpty
            cruzamento33.d21094 = vbEmpty
            cruzamento33.d10279 = vbEmpty
            cruzamento33.d10280 = vbEmpty
            cruzamento33.d10910 = vbEmpty
            cruzamento33.d14123 = vbEmpty
            cruzamento33.top5 = vbEmpty
            cruzamento33.dci_obr = vbEmpty
            cruzamento33.lei6 = vbEmpty
            cruzamento33.pvpmenos1 = vbEmpty
            cruzamento33.pvpmenos2 = vbEmpty
            cruzamento33.trocamarca = vbEmpty
            cruzamento33.CNPEM = vbEmpty
            cruzamento33.nDCI = vbEmpty  '?
            cruzamento33.excepa = vbEmpty 'se 
            cruzamento33.excepb = vbEmpty
            cruzamento33.excepc = vbEmpty
            cruzamento33.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento33.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento33.xgh = vbEmpty 'se existem ambos
            cruzamento33.xcnpem = vbEmpty 'se existem ambos
            cruzamento33.xnDCI = vbEmpty 'se existem ambos
            cruzamento33.porDCImesmoCNPEM = vbEmpty
            cruzamento33.porDCIdifCNPEM = vbEmpty
            cruzamento33.porMARCAmesmoCNPEM = vbEmpty
            cruzamento33.porMARCAdifCNPEM = vbEmpty
            cruzamento33.ranking = vbEmpty
            cruzamento33.anulado = vbEmpty
            cruzamento33.qualP = vbEmpty
            cruzamento33.marcadifsemhavergens = vbEmpty

            cruzamento34.code = vbEmpty
            cruzamento34.dci = vbEmpty
            cruzamento34.nome = vbEmpty
            cruzamento34.forma = vbEmpty
            cruzamento34.dose = vbEmpty
            cruzamento34.qty = vbEmpty
            cruzamento34.comp = vbEmpty
            cruzamento34.gh = vbEmpty
            cruzamento34.pvp = vbEmpty
            cruzamento34.pr = vbEmpty
            cruzamento34.gen = vbEmpty
            cruzamento34.lab = vbEmpty
            cruzamento34.d4250 = vbEmpty
            cruzamento34.d1234 = vbEmpty
            cruzamento34.d21094 = vbEmpty
            cruzamento34.d10279 = vbEmpty
            cruzamento34.d10280 = vbEmpty
            cruzamento34.d10910 = vbEmpty
            cruzamento34.d14123 = vbEmpty
            cruzamento34.top5 = vbEmpty
            cruzamento34.dci_obr = vbEmpty
            cruzamento34.lei6 = vbEmpty
            cruzamento34.pvpmenos1 = vbEmpty
            cruzamento34.pvpmenos2 = vbEmpty
            cruzamento34.trocamarca = vbEmpty
            cruzamento34.CNPEM = vbEmpty
            cruzamento34.nDCI = vbEmpty  '?
            cruzamento34.excepa = vbEmpty 'se 
            cruzamento34.excepb = vbEmpty
            cruzamento34.excepc = vbEmpty
            cruzamento34.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento34.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento34.xgh = vbEmpty 'se existem ambos
            cruzamento34.xcnpem = vbEmpty 'se existem ambos
            cruzamento34.xnDCI = vbEmpty 'se existem ambos
            cruzamento34.porDCImesmoCNPEM = vbEmpty
            cruzamento34.porDCIdifCNPEM = vbEmpty
            cruzamento34.porMARCAmesmoCNPEM = vbEmpty
            cruzamento34.porMARCAdifCNPEM = vbEmpty
            cruzamento34.ranking = vbEmpty
            cruzamento34.anulado = vbEmpty
            cruzamento34.qualP = vbEmpty
            cruzamento34.marcadifsemhavergens = vbEmpty


            cruzamento21.code = vbEmpty
            cruzamento21.dci = vbEmpty
            cruzamento21.nome = vbEmpty
            cruzamento21.forma = vbEmpty
            cruzamento21.dose = vbEmpty
            cruzamento21.qty = vbEmpty
            cruzamento21.comp = vbEmpty
            cruzamento21.gh = vbEmpty
            cruzamento21.pvp = vbEmpty
            cruzamento21.pr = vbEmpty
            cruzamento21.gen = vbEmpty
            cruzamento21.lab = vbEmpty
            cruzamento21.d4250 = vbEmpty
            cruzamento21.d1234 = vbEmpty
            cruzamento21.d21094 = vbEmpty
            cruzamento21.d10279 = vbEmpty
            cruzamento21.d10280 = vbEmpty
            cruzamento21.d10910 = vbEmpty
            cruzamento21.d14123 = vbEmpty
            cruzamento21.top5 = vbEmpty
            cruzamento21.dci_obr = vbEmpty
            cruzamento21.lei6 = vbEmpty
            cruzamento21.pvpmenos1 = vbEmpty
            cruzamento21.pvpmenos2 = vbEmpty
            cruzamento21.trocamarca = vbEmpty
            cruzamento21.CNPEM = vbEmpty
            cruzamento21.nDCI = vbEmpty  '?
            cruzamento21.excepa = vbEmpty 'se 
            cruzamento21.excepb = vbEmpty
            cruzamento21.excepc = vbEmpty
            cruzamento21.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento21.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento21.xgh = vbEmpty 'se existem ambos
            cruzamento21.xcnpem = vbEmpty 'se existem ambos
            cruzamento21.xnDCI = vbEmpty 'se existem ambos
            cruzamento21.porDCImesmoCNPEM = vbEmpty
            cruzamento21.porDCIdifCNPEM = vbEmpty
            cruzamento21.porMARCAmesmoCNPEM = vbEmpty
            cruzamento21.porMARCAdifCNPEM = vbEmpty
            cruzamento21.ranking = vbEmpty
            cruzamento21.anulado = vbEmpty
            cruzamento21.qualP = vbEmpty
            cruzamento21.marcadifsemhavergens = vbEmpty
            cruzamento22.code = vbEmpty
            cruzamento22.dci = vbEmpty
            cruzamento22.nome = vbEmpty
            cruzamento22.forma = vbEmpty
            cruzamento22.dose = vbEmpty
            cruzamento22.qty = vbEmpty
            cruzamento22.comp = vbEmpty
            cruzamento22.gh = vbEmpty
            cruzamento22.pvp = vbEmpty
            cruzamento22.pr = vbEmpty
            cruzamento22.gen = vbEmpty
            cruzamento22.lab = vbEmpty
            cruzamento22.d4250 = vbEmpty
            cruzamento22.d1234 = vbEmpty
            cruzamento22.d21094 = vbEmpty
            cruzamento22.d10279 = vbEmpty
            cruzamento22.d10280 = vbEmpty
            cruzamento22.d10910 = vbEmpty
            cruzamento22.d14123 = vbEmpty
            cruzamento22.top5 = vbEmpty
            cruzamento22.dci_obr = vbEmpty
            cruzamento22.lei6 = vbEmpty
            cruzamento22.pvpmenos1 = vbEmpty
            cruzamento22.pvpmenos2 = vbEmpty
            cruzamento22.trocamarca = vbEmpty
            cruzamento22.CNPEM = vbEmpty
            cruzamento22.nDCI = vbEmpty  '?
            cruzamento22.excepa = vbEmpty 'se 
            cruzamento22.excepb = vbEmpty
            cruzamento22.excepc = vbEmpty
            cruzamento22.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento22.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento22.xgh = vbEmpty 'se existem ambos
            cruzamento22.xcnpem = vbEmpty 'se existem ambos
            cruzamento22.xnDCI = vbEmpty 'se existem ambos
            cruzamento22.porDCImesmoCNPEM = vbEmpty
            cruzamento22.porDCIdifCNPEM = vbEmpty
            cruzamento22.porMARCAmesmoCNPEM = vbEmpty
            cruzamento22.porMARCAdifCNPEM = vbEmpty
            cruzamento22.ranking = vbEmpty
            cruzamento22.anulado = vbEmpty
            cruzamento22.qualP = vbEmpty
            cruzamento22.marcadifsemhavergens = vbEmpty

            cruzamento23.code = vbEmpty
            cruzamento23.dci = vbEmpty
            cruzamento23.nome = vbEmpty
            cruzamento23.forma = vbEmpty
            cruzamento23.dose = vbEmpty
            cruzamento23.qty = vbEmpty
            cruzamento23.comp = vbEmpty
            cruzamento23.gh = vbEmpty
            cruzamento23.pvp = vbEmpty
            cruzamento23.pr = vbEmpty
            cruzamento23.gen = vbEmpty
            cruzamento23.lab = vbEmpty
            cruzamento23.d4250 = vbEmpty
            cruzamento23.d1234 = vbEmpty
            cruzamento23.d21094 = vbEmpty
            cruzamento23.d10279 = vbEmpty
            cruzamento23.d10280 = vbEmpty
            cruzamento23.d10910 = vbEmpty
            cruzamento23.d14123 = vbEmpty
            cruzamento23.top5 = vbEmpty
            cruzamento23.dci_obr = vbEmpty
            cruzamento23.lei6 = vbEmpty
            cruzamento23.pvpmenos1 = vbEmpty
            cruzamento23.pvpmenos2 = vbEmpty
            cruzamento23.trocamarca = vbEmpty
            cruzamento23.CNPEM = vbEmpty
            cruzamento23.nDCI = vbEmpty  '?
            cruzamento23.excepa = vbEmpty 'se 
            cruzamento23.excepb = vbEmpty
            cruzamento23.excepc = vbEmpty
            cruzamento23.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento23.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento23.xgh = vbEmpty 'se existem ambos
            cruzamento23.xcnpem = vbEmpty 'se existem ambos
            cruzamento23.xnDCI = vbEmpty 'se existem ambos
            cruzamento23.porDCImesmoCNPEM = vbEmpty
            cruzamento23.porDCIdifCNPEM = vbEmpty
            cruzamento23.porMARCAmesmoCNPEM = vbEmpty
            cruzamento23.porMARCAdifCNPEM = vbEmpty
            cruzamento23.ranking = vbEmpty
            cruzamento23.anulado = vbEmpty
            cruzamento23.qualP = vbEmpty
            cruzamento23.marcadifsemhavergens = vbEmpty

            cruzamento24.code = vbEmpty
            cruzamento24.dci = vbEmpty
            cruzamento24.nome = vbEmpty
            cruzamento24.forma = vbEmpty
            cruzamento24.dose = vbEmpty
            cruzamento24.qty = vbEmpty
            cruzamento24.comp = vbEmpty
            cruzamento24.gh = vbEmpty
            cruzamento24.pvp = vbEmpty
            cruzamento24.pr = vbEmpty
            cruzamento24.gen = vbEmpty
            cruzamento24.lab = vbEmpty
            cruzamento24.d4250 = vbEmpty
            cruzamento24.d1234 = vbEmpty
            cruzamento24.d21094 = vbEmpty
            cruzamento24.d10279 = vbEmpty
            cruzamento24.d10280 = vbEmpty
            cruzamento24.d10910 = vbEmpty
            cruzamento24.d14123 = vbEmpty
            cruzamento24.top5 = vbEmpty
            cruzamento24.dci_obr = vbEmpty
            cruzamento24.lei6 = vbEmpty
            cruzamento24.pvpmenos1 = vbEmpty
            cruzamento24.pvpmenos2 = vbEmpty
            cruzamento24.trocamarca = vbEmpty
            cruzamento24.CNPEM = vbEmpty
            cruzamento24.nDCI = vbEmpty  '?
            cruzamento24.excepa = vbEmpty 'se 
            cruzamento24.excepb = vbEmpty
            cruzamento24.excepc = vbEmpty
            cruzamento24.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento24.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento24.xgh = vbEmpty 'se existem ambos
            cruzamento24.xcnpem = vbEmpty 'se existem ambos
            cruzamento24.xnDCI = vbEmpty 'se existem ambos
            cruzamento24.porDCImesmoCNPEM = vbEmpty
            cruzamento24.porDCIdifCNPEM = vbEmpty
            cruzamento24.porMARCAmesmoCNPEM = vbEmpty
            cruzamento24.porMARCAdifCNPEM = vbEmpty
            cruzamento24.ranking = vbEmpty
            cruzamento24.anulado = vbEmpty
            cruzamento24.qualP = vbEmpty
            cruzamento24.marcadifsemhavergens = vbEmpty

            cruzamento11.code = vbEmpty
            cruzamento11.dci = vbEmpty
            cruzamento11.nome = vbEmpty
            cruzamento11.forma = vbEmpty
            cruzamento11.dose = vbEmpty
            cruzamento11.qty = vbEmpty
            cruzamento11.comp = vbEmpty
            cruzamento11.gh = vbEmpty
            cruzamento11.pvp = vbEmpty
            cruzamento11.pr = vbEmpty
            cruzamento11.gen = vbEmpty
            cruzamento11.lab = vbEmpty
            cruzamento11.d4250 = vbEmpty
            cruzamento11.d1234 = vbEmpty
            cruzamento11.d21094 = vbEmpty
            cruzamento11.d10279 = vbEmpty
            cruzamento11.d10280 = vbEmpty
            cruzamento11.d10910 = vbEmpty
            cruzamento11.d14123 = vbEmpty
            cruzamento11.top5 = vbEmpty
            cruzamento11.dci_obr = vbEmpty
            cruzamento11.lei6 = vbEmpty
            cruzamento11.pvpmenos1 = vbEmpty
            cruzamento11.pvpmenos2 = vbEmpty
            cruzamento11.trocamarca = vbEmpty
            cruzamento11.CNPEM = vbEmpty
            cruzamento11.nDCI = vbEmpty  '?
            cruzamento11.excepa = vbEmpty 'se 
            cruzamento11.excepb = vbEmpty
            cruzamento11.excepc = vbEmpty
            cruzamento11.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento11.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento11.xgh = vbEmpty 'se existem ambos
            cruzamento11.xcnpem = vbEmpty 'se existem ambos
            cruzamento11.xnDCI = vbEmpty 'se existem ambos
            cruzamento11.porDCImesmoCNPEM = vbEmpty
            cruzamento11.porDCIdifCNPEM = vbEmpty
            cruzamento11.porMARCAmesmoCNPEM = vbEmpty
            cruzamento11.porMARCAdifCNPEM = vbEmpty
            cruzamento11.ranking = vbEmpty
            cruzamento11.anulado = vbEmpty
            cruzamento11.qualP = vbEmpty
            cruzamento11.marcadifsemhavergens = vbEmpty
            cruzamento12.code = vbEmpty
            cruzamento12.dci = vbEmpty
            cruzamento12.nome = vbEmpty
            cruzamento12.forma = vbEmpty
            cruzamento12.dose = vbEmpty
            cruzamento12.qty = vbEmpty
            cruzamento12.comp = vbEmpty
            cruzamento12.gh = vbEmpty
            cruzamento12.pvp = vbEmpty
            cruzamento12.pr = vbEmpty
            cruzamento12.gen = vbEmpty
            cruzamento12.lab = vbEmpty
            cruzamento12.d4250 = vbEmpty
            cruzamento12.d1234 = vbEmpty
            cruzamento12.d21094 = vbEmpty
            cruzamento12.d10279 = vbEmpty
            cruzamento12.d10280 = vbEmpty
            cruzamento12.d10910 = vbEmpty
            cruzamento12.d14123 = vbEmpty
            cruzamento12.top5 = vbEmpty
            cruzamento12.dci_obr = vbEmpty
            cruzamento12.lei6 = vbEmpty
            cruzamento12.pvpmenos1 = vbEmpty
            cruzamento12.pvpmenos2 = vbEmpty
            cruzamento12.trocamarca = vbEmpty
            cruzamento12.CNPEM = vbEmpty
            cruzamento12.nDCI = vbEmpty  '?
            cruzamento12.excepa = vbEmpty 'se 
            cruzamento12.excepb = vbEmpty
            cruzamento12.excepc = vbEmpty
            cruzamento12.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento12.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento12.xgh = vbEmpty 'se existem ambos
            cruzamento12.xcnpem = vbEmpty 'se existem ambos
            cruzamento12.xnDCI = vbEmpty 'se existem ambos
            cruzamento12.porDCImesmoCNPEM = vbEmpty
            cruzamento12.porDCIdifCNPEM = vbEmpty
            cruzamento12.porMARCAmesmoCNPEM = vbEmpty
            cruzamento12.porMARCAdifCNPEM = vbEmpty
            cruzamento12.ranking = vbEmpty
            cruzamento12.anulado = vbEmpty
            cruzamento12.qualP = vbEmpty
            cruzamento12.marcadifsemhavergens = vbEmpty

            cruzamento13.code = vbEmpty
            cruzamento13.dci = vbEmpty
            cruzamento13.nome = vbEmpty
            cruzamento13.forma = vbEmpty
            cruzamento13.dose = vbEmpty
            cruzamento13.qty = vbEmpty
            cruzamento13.comp = vbEmpty
            cruzamento13.gh = vbEmpty
            cruzamento13.pvp = vbEmpty
            cruzamento13.pr = vbEmpty
            cruzamento13.gen = vbEmpty
            cruzamento13.lab = vbEmpty
            cruzamento13.d4250 = vbEmpty
            cruzamento13.d1234 = vbEmpty
            cruzamento13.d21094 = vbEmpty
            cruzamento13.d10279 = vbEmpty
            cruzamento13.d10280 = vbEmpty
            cruzamento13.d10910 = vbEmpty
            cruzamento13.d14123 = vbEmpty
            cruzamento13.top5 = vbEmpty
            cruzamento13.dci_obr = vbEmpty
            cruzamento13.lei6 = vbEmpty
            cruzamento13.pvpmenos1 = vbEmpty
            cruzamento13.pvpmenos2 = vbEmpty
            cruzamento13.trocamarca = vbEmpty
            cruzamento13.CNPEM = vbEmpty
            cruzamento13.nDCI = vbEmpty  '?
            cruzamento13.excepa = vbEmpty 'se 
            cruzamento13.excepb = vbEmpty
            cruzamento13.excepc = vbEmpty
            cruzamento13.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento13.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento13.xgh = vbEmpty 'se existem ambos
            cruzamento13.xcnpem = vbEmpty 'se existem ambos
            cruzamento13.xnDCI = vbEmpty 'se existem ambos
            cruzamento13.porDCImesmoCNPEM = vbEmpty
            cruzamento13.porDCIdifCNPEM = vbEmpty
            cruzamento13.porMARCAmesmoCNPEM = vbEmpty
            cruzamento13.porMARCAdifCNPEM = vbEmpty
            cruzamento13.ranking = vbEmpty
            cruzamento13.anulado = vbEmpty
            cruzamento13.qualP = vbEmpty
            cruzamento13.marcadifsemhavergens = vbEmpty

            cruzamento14.code = vbEmpty
            cruzamento14.dci = vbEmpty
            cruzamento14.nome = vbEmpty
            cruzamento14.forma = vbEmpty
            cruzamento14.dose = vbEmpty
            cruzamento14.qty = vbEmpty
            cruzamento14.comp = vbEmpty
            cruzamento14.gh = vbEmpty
            cruzamento14.pvp = vbEmpty
            cruzamento14.pr = vbEmpty
            cruzamento14.gen = vbEmpty
            cruzamento14.lab = vbEmpty
            cruzamento14.d4250 = vbEmpty
            cruzamento14.d1234 = vbEmpty
            cruzamento14.d21094 = vbEmpty
            cruzamento14.d10279 = vbEmpty
            cruzamento14.d10280 = vbEmpty
            cruzamento14.d10910 = vbEmpty
            cruzamento14.d14123 = vbEmpty
            cruzamento14.top5 = vbEmpty
            cruzamento14.dci_obr = vbEmpty
            cruzamento14.lei6 = vbEmpty
            cruzamento14.pvpmenos1 = vbEmpty
            cruzamento14.pvpmenos2 = vbEmpty
            cruzamento14.trocamarca = vbEmpty
            cruzamento14.CNPEM = vbEmpty
            cruzamento14.nDCI = vbEmpty  '?
            cruzamento14.excepa = vbEmpty 'se 
            cruzamento14.excepb = vbEmpty
            cruzamento14.excepc = vbEmpty
            cruzamento14.xqty = vbEmpty 'menor, 1<x<=1,5, x>1,5
            cruzamento14.xpvp = vbEmpty 'se existem ambos e compara maior, menor
            cruzamento14.xgh = vbEmpty 'se existem ambos
            cruzamento14.xcnpem = vbEmpty 'se existem ambos
            cruzamento14.xnDCI = vbEmpty 'se existem ambos
            cruzamento14.porDCImesmoCNPEM = vbEmpty
            cruzamento14.porDCIdifCNPEM = vbEmpty
            cruzamento14.porMARCAmesmoCNPEM = vbEmpty
            cruzamento14.porMARCAdifCNPEM = vbEmpty
            cruzamento14.ranking = vbEmpty
            cruzamento14.anulado = vbEmpty
            cruzamento14.qualP = vbEmpty
            cruzamento14.marcadifsemhavergens = vbEmpty
            resultado1.qualP = 0
            resultado2.qualP = 0
            resultado3.qualP = 0
            resultado4.qualP = 0
            resultado1.ranking = "000000"
            resultado2.ranking = "000000"
            resultado3.ranking = "000000"
            resultado4.ranking = "000000"
            cruzamento11.anulado = False
            cruzamento12.anulado = False
            cruzamento13.anulado = False
            cruzamento14.anulado = False
            cruzamento21.anulado = False
            cruzamento22.anulado = False
            cruzamento23.anulado = False
            cruzamento24.anulado = False
            cruzamento31.anulado = False
            cruzamento32.anulado = False
            cruzamento33.anulado = False
            cruzamento34.anulado = False
            cruzamento41.anulado = False
            cruzamento42.anulado = False
            cruzamento43.anulado = False
            cruzamento44.anulado = False
            cruzamento11.ranking = "0000"
            cruzamento12.ranking = "0000"
            cruzamento13.ranking = "0000"
            cruzamento14.ranking = "0000"
            cruzamento21.ranking = "0000"
            cruzamento22.ranking = "0000"
            cruzamento23.ranking = "0000"
            cruzamento24.ranking = "0000"
            cruzamento31.ranking = "0000"
            cruzamento32.ranking = "0000"
            cruzamento33.ranking = "0000"
            cruzamento34.ranking = "0000"
            cruzamento41.ranking = "0000"
            cruzamento42.ranking = "0000"
            cruzamento43.ranking = "0000"
            cruzamento44.ranking = "0000"

            op4.code = vbEmpty
            op4.dci = vbEmpty
            op4.forma = vbEmpty
            op4.dose = vbEmpty
            op4.qty = vbEmpty
            op4.comp = vbEmpty
            op4.gh = vbEmpty
            op4.gen = vbEmpty
            op4.lab = vbEmpty
            op4.pvp = vbEmpty
            op4.pr = vbEmpty
            op4.d4250 = vbEmpty
            op4.d1234 = vbEmpty
            op4.d10279 = vbEmpty
            op4.d10280 = vbEmpty
            op4.d21094 = vbEmpty
            op4.lei6 = vbEmpty
            op4.d14123 = vbEmpty
            op4.dci_obr = vbEmpty
            op4.trocamarca = vbEmpty
            op4.top5 = vbEmpty
            op4.pvpmenos1 = vbEmpty
            op4.CNPEM = vbEmpty
            op4.porCNPEM = vbEmpty
            op4.duplicado = vbEmpty
            op4.nDCI = vbEmpty
            op4.pvpmenos2 = vbEmpty
            op4.pvptop5 = vbEmpty
            op4.pvpmenos1top5 = vbEmpty
            op4.pvpmenos2top5 = vbEmpty
            op4.pvpun = vbEmpty
            op4.pvpmenos1un = vbEmpty
            op4.pvpmenos2un = vbEmpty
            op4.excepa = vbEmpty
            op4.excepb = vbEmpty
            op4.prescrito = vbEmpty
            op4.cruzacom = vbEmpty
            op4.result = vbEmpty
            op3.code = vbEmpty
            op3.dci = vbEmpty
            op3.forma = vbEmpty
            op3.dose = vbEmpty
            op3.qty = vbEmpty
            op3.comp = vbEmpty
            op3.gh = vbEmpty
            op3.gen = vbEmpty
            op3.lab = vbEmpty
            op3.pvp = vbEmpty
            op3.pr = vbEmpty
            op3.d4250 = vbEmpty
            op3.d1234 = vbEmpty
            op3.d10279 = vbEmpty
            op3.d10280 = vbEmpty
            op3.d21094 = vbEmpty
            op3.lei6 = vbEmpty
            op3.d14123 = vbEmpty
            op3.dci_obr = vbEmpty
            op3.trocamarca = vbEmpty
            op3.top5 = vbEmpty
            op3.pvpmenos1 = vbEmpty
            op3.CNPEM = vbEmpty
            op3.porCNPEM = vbEmpty
            op3.duplicado = vbEmpty
            op3.nDCI = vbEmpty
            op3.pvpmenos2 = vbEmpty
            op3.pvptop5 = vbEmpty
            op3.pvpmenos1top5 = vbEmpty
            op3.pvpmenos2top5 = vbEmpty
            op3.pvpun = vbEmpty
            op3.pvpmenos1un = vbEmpty
            op3.pvpmenos2un = vbEmpty
            op3.excepa = vbEmpty
            op3.excepb = vbEmpty
            op3.prescrito = vbEmpty
            op3.cruzacom = vbEmpty
            op3.result = vbEmpty
            op2.code = vbEmpty
            op2.dci = vbEmpty
            op2.forma = vbEmpty
            op2.dose = vbEmpty
            op2.qty = vbEmpty
            op2.comp = vbEmpty
            op2.gh = vbEmpty
            op2.gen = vbEmpty
            op2.lab = vbEmpty
            op2.pvp = vbEmpty
            op2.pr = vbEmpty
            op2.d4250 = vbEmpty
            op2.d1234 = vbEmpty
            op2.d10279 = vbEmpty
            op2.d10280 = vbEmpty
            op2.d21094 = vbEmpty
            op2.lei6 = vbEmpty
            op2.d14123 = vbEmpty
            op2.dci_obr = vbEmpty
            op2.trocamarca = vbEmpty
            op2.top5 = vbEmpty
            op2.pvpmenos1 = vbEmpty
            op2.CNPEM = vbEmpty
            op2.porCNPEM = vbEmpty
            op2.duplicado = vbEmpty
            op2.nDCI = vbEmpty
            op2.pvpmenos2 = vbEmpty
            op2.pvptop5 = vbEmpty
            op2.pvpmenos1top5 = vbEmpty
            op2.pvpmenos2top5 = vbEmpty
            op2.pvpun = vbEmpty
            op2.pvpmenos1un = vbEmpty
            op2.pvpmenos2un = vbEmpty
            op2.excepa = vbEmpty
            op2.excepb = vbEmpty
            op2.prescrito = vbEmpty
            op2.cruzacom = vbEmpty
            op2.result = vbEmpty
            op1.code = vbEmpty
            op1.dci = vbEmpty
            op1.forma = vbEmpty
            op1.dose = vbEmpty
            op1.qty = vbEmpty
            op1.comp = vbEmpty
            op1.gh = vbEmpty
            op1.gen = vbEmpty
            op1.lab = vbEmpty
            op1.pvp = vbEmpty
            op1.pr = vbEmpty
            op1.d4250 = vbEmpty
            op1.d1234 = vbEmpty
            op1.d10279 = vbEmpty
            op1.d10280 = vbEmpty
            op1.d21094 = vbEmpty
            op1.lei6 = vbEmpty
            op1.d14123 = vbEmpty
            op1.dci_obr = vbEmpty
            op1.trocamarca = vbEmpty
            op1.top5 = vbEmpty
            op1.pvpmenos1 = vbEmpty
            op1.CNPEM = vbEmpty
            op1.porCNPEM = vbEmpty
            op1.duplicado = vbEmpty
            op1.nDCI = vbEmpty
            op1.pvpmenos2 = vbEmpty
            op1.pvptop5 = vbEmpty
            op1.pvpmenos1top5 = vbEmpty
            op1.pvpmenos2top5 = vbEmpty
            op1.pvpun = vbEmpty
            op1.pvpmenos1un = vbEmpty
            op1.pvpmenos2un = vbEmpty
            op1.excepa = vbEmpty
            op1.excepb = vbEmpty
            op1.prescrito = vbEmpty
            op1.cruzacom = vbEmpty
            op1.result = vbEmpty
            vshe = 0
            vshe1 = 0
            vshe2 = 0
            vshe3 = 0
            vshe4 = 0
            limpo = True
            agrupado = False


a:          varlabelcruz1 = ""
b:
1:          av1.nivel = 100
2:          av2.nivel = 100
3:          av3.nivel = 100
4:          av4.nivel = 100
5:          a1_4250 = False
6:          a2_4250 = False
7:          a3_4250 = False
8:          a4_4250 = False
9:          a1_1234 = False
10:         a2_1234 = False
11:         a3_1234 = False
12:         a4_1234 = False
13:         a1_14123 = False
14:         a2_14123 = False
15:         a3_14123 = False
16:         a4_14123 = False
17:         a1_21094 = False
18:         a2_21094 = False
19:         a3_21094 = False
20:         a4_21094 = False
21:         a1_1474ad = False
22:         a2_1474ad = False
23:         a3_1474ad = False
24:         a4_1474ad = False
25:         a1_1474nl = False
26:         a2_1474nl = False
27:         a3_1474nl = False
28:         a4_1474nl = False
29:         a1_10279 = False
30:         a2_10279 = False
31:         a3_10279 = False
32:         a4_10279 = False
33:         a1_10910 = False
34:         a2_10910 = False
35:         a3_10910 = False
36:         a4_10910 = False
37:         descomp1mostrado = False
38:         descomp2mostrado = False
39:         descomp3mostrado = False
40:         descomp4mostrado = False
41:         A = 0
42:         P = 0
43:
47:         Me.presc1.BackColor = Color.White
48:         Me.presc2.BackColor = Color.White
49:         Me.presc3.BackColor = Color.White
50:         Me.presc4.BackColor = Color.White
51:         Me.aviam1.BackColor = Color.White
52:         Me.aviam2.BackColor = Color.White
53:         Me.aviam3.BackColor = Color.White
54:         Me.aviam4.BackColor = Color.White
55:         Me.aviam1.Text = ""
56:         Me.aviam2.Text = ""
57:         Me.aviam3.Text = ""
58:         Me.aviam4.Text = ""
59:         Me.presc1.Text = ""
60:         Me.presc2.Text = ""
61:         Me.presc3.Text = ""
62:         Me.presc4.Text = ""
63:         Me.result1.Text = ""
64:         Me.result2.Text = ""
65:         Me.result3.Text = ""
66:         Me.result4.Text = ""
67:         Me.result1.BackColor = cordefundo
68:         Me.result2.BackColor = cordefundo
69:         Me.result3.BackColor = cordefundo
70:         Me.result4.BackColor = cordefundo
71:         Me.presc1.Focus()
72:         Me.Prescrito1.codigo = Nothing
73:         Me.Prescrito2.codigo = Nothing
74:         Me.Prescrito3.codigo = Nothing
75:         Me.Prescrito4.codigo = Nothing
76:         Me.Aviado1.codigo = Nothing
77:         Me.Aviado2.codigo = Nothing
78:         Me.Aviado3.codigo = Nothing
79:         Me.Aviado4.codigo = Nothing
80:         Me.a1p1.nivel = 100
81:         Me.a2p1.nivel = 100
82:         Me.a3p1.nivel = 100
83:         Me.a4p1.nivel = 100
84:         Me.a1p2.nivel = 100
85:         Me.a2p2.nivel = 100
86:         Me.a3p2.nivel = 100
87:         Me.a4p2.nivel = 100
88:         Me.a1p3.nivel = 100
89:         Me.a2p3.nivel = 100
90:         Me.a3p3.nivel = 100
91:         Me.a4p3.nivel = 100
92:         Me.a1p4.nivel = 100
93:         Me.a2p4.nivel = 100
94:         Me.a3p4.nivel = 100
95:         Me.a4p4.nivel = 100

96:         Me.a1p1.resultado = 100
97:         Me.a2p1.resultado = 100
98:         Me.a3p1.resultado = 100
99:         Me.a4p1.resultado = 100
100:        Me.a1p2.resultado = 100
101:        Me.a2p2.resultado = 100
102:        Me.a3p2.resultado = 100
103:        Me.a4p2.resultado = 100
104:        Me.a1p3.resultado = 100
105:        Me.a2p3.resultado = 100
106:        Me.a3p3.resultado = 100
107:        Me.a4p3.resultado = 100
108:        Me.a1p4.resultado = 100
109:        Me.a2p4.resultado = 100
110:        Me.a3p4.resultado = 100
111:        Me.a4p4.resultado = 100
112:        Me.a1p1.mostrado = False
113:        Me.a2p1.mostrado = False
114:        Me.a3p1.mostrado = False
115:        Me.a4p1.mostrado = False
116:        Me.a1p2.mostrado = False
117:        Me.a2p2.mostrado = False
118:        Me.a3p2.mostrado = False
119:        Me.a4p2.mostrado = False
120:        Me.a1p3.mostrado = False
121:        Me.a2p3.mostrado = False
122:        Me.a3p3.mostrado = False
123:        Me.a4p3.mostrado = False
124:        Me.a1p4.mostrado = False
125:        Me.a2p4.mostrado = False
126:        Me.a3p4.mostrado = False
127:        Me.a4p4.mostrado = False
128:
136:        grupoP1 = 0
137:        grupoP2 = 0
138:        grupoP3 = 0
139:        grupoP4 = 0
140:        grupoA1 = 0
141:        grupoA2 = 0
142:        grupoA3 = 0
143:        grupoA4 = 0
144:        grupoP1dci = ""
145:        grupoP2dci = ""
146:        grupoP3dci = ""
147:        grupoP4dci = ""
148:        grupoA1dci = ""
149:        grupoA2dci = ""
150:        grupoA3dci = ""
151:        grupoA4dci = ""
152:        varcruza1 = 0
153:        varcruza2 = 0
154:        varcruza3 = 0
155:        varcruza4 = 0
156:        'varcruza5 = 0
157:        'varcruza6 = 0
158:        varcruzp1 = 0
159:        varcruzp2 = 0
160:        varcruzp3 = 0
161:        varcruzp4 = 0
162:
178:
195:        verde = True
196:        amarelo = False
197:        vermelho = False
198:
            conjunto = 0
            inicializado = True
        End If

        Exit Sub
MOSTRARERRO:
        MsgBox("SUB INICIALIZAR: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'lança a função comparar ao fazer enter no botão com o mesmo nome
    Private Sub ButComparar_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      Comparar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub ButComparar_Enter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'lança a função comparar ao clickar no botão com o mesmo nome
    Private Sub ButComparar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      Comparar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub ButComparar_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Dim bdpath As String
    'o que acontece ao abrir/inicializar o form principal - inicia timer, limpa tudo, poe tudo a zero e foca na 1ª textbox
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'BasededadosDataSet1.infarmed' table. You can move, or remove it, as needed.
        On Error GoTo MOSTRARERRO
        mesactual = DateTime.Today.Month
        bdpath = "basededados.mdb"
        PC.nome = My.Computer.Name
        PC.sistema = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        PC.fileinicial = PC.sistema & "\ofpafile."
        Me.Text = Application.ProductName & " " & Application.ProductVersion & " (" & FileDateTime(bdpath).ToString.Substring(6, 4) & FileDateTime(bdpath).ToString.Substring(3, 2) & FileDateTime(bdpath).ToString.Substring(0, 2) & ")"
        rede = FileSystem.CurDir
        progressao.Show()
2:      If autenticacao() = False Then
            Me.Close()
        End If
        month()

3:      Me.InfarmedTableAdapter1.Fill(Me.BasededadosDataSet1.infarmed)
        progressao.Hide()
4:      'Me.Width = 1000
5:      'Me.Height = 640
6:      Timer1.Start()
        showcode = True
7:      inicializar()
        inicializado = False
8:      Me.WindowState = FormWindowState.Normal
9:      Me.KeyPreview = True

10:     Exit Sub
MOSTRARERRO:
        MsgBox("Sub Form1_Load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub data_Keyspace(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'estava a repetir por isso tirei daqui e pus no Keys_enter
        'If e.KeyCode = Keys.Space Then
        'e.SuppressKeyPress = True

        'inicializar()

178:    'Me.Focus()

179:    'inicializar()
        'Me.presc1.Text = ""
        'End If
    End Sub


    'os próximos 7 são para não aceitar foco (focado pelo utilizador) a não ser na 1ª textbox livre
    Private Sub focadop2(sender As Object, e As System.EventArgs) Handles presc2.GotFocus
        If presc1.Text = "" Then
            presc1.Focus()
        End If
    End Sub
    Private Sub focadop3(sender As Object, e As System.EventArgs) Handles presc3.GotFocus
        If presc2.Text = "" Then
            presc2.Focus()
        End If
    End Sub
    Private Sub focadop4(sender As Object, e As System.EventArgs) Handles presc4.GotFocus
        If presc3.Text = "" Then
            presc3.Focus()
        End If

    End Sub
    Private Sub focadoa1(sender As Object, e As System.EventArgs) Handles aviam1.GotFocus
        If presc1.Text = "" Then
            presc1.Focus()
        End If

    End Sub
    Private Sub focadoa2(sender As Object, e As System.EventArgs) Handles aviam2.GotFocus
        If aviam1.Text = "" Then
            aviam1.Focus()
        End If

    End Sub
    Private Sub focadoa3(sender As Object, e As System.EventArgs) Handles aviam3.GotFocus
        If aviam2.Text = "" Then
            aviam2.Focus()
        End If
    End Sub
    Private Sub focadoa4(sender As Object, e As System.EventArgs) Handles aviam4.GotFocus
        If aviam3.Text = "" Then
            aviam3.Focus()
        End If
    End Sub



    Private Sub data_KeysEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        inicializado = False
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True

            inicializar()

178:        Me.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Select Case foco

                Case "presc1"
55:                 If presc1.Text = "" Then
56:                     'Beep()
57:                 Else
                        If IsNumeric(presc1.Text) Then
58:                         Me.presc2.Focus()
                        Else
                            presc1.Text = ""
                        End If
59:                 End If
60:
61:             Case "aviam1"
62:                 If aviam1.Text = "" Then
63:                     'Beep()
64:                 Else
                        If IsNumeric(aviam1.Text) Then
                            If aviam1.Text >= 1111111 And aviam1.Text <= 9999999 Then
65:                             Me.aviam2.Focus()
                            End If
                        Else
                            aviam1.Text = ""
                        End If
66:                 End If
67:             Case "presc2"
68:                 If presc2.Text = "" Then
69:                     Me.aviam1.Focus()
70:                     Me.presc2.Text = "0"
71:                     Me.presc3.Text = "0"
72:                     Me.presc4.Text = "0"
a72:                    'Me.presc5.Text = "0"
b72:                    'Me.presc6.Text = "0"
73:                 Else
                        If IsNumeric(presc2.Text) Then
74:                         Me.presc3.Focus()
                        Else
                            presc2.Text = ""
                        End If
75:                 End If
76:             Case "presc3"
77:                 If presc3.Text = "" Then
78:                     Me.aviam1.Focus()
79:                     Me.presc3.Text = "0"
80:                     Me.presc4.Text = "0"
a80:                    'Me.presc5.Text = "0"
b80:                    'Me.presc6.Text = "0"
81:                 Else
                        If IsNumeric(presc3.Text) Then
82:                         Me.presc4.Focus()
                        Else
                            presc3.Text = ""
                        End If
83:                 End If
84:             Case "presc4"
85:                 If presc4.Text = "" Then
86:                     Me.aviam1.Focus()
87:                     Me.presc4.Text = "0"
a87:                    'Me.presc5.Text = "0"
b87:                    'Me.presc6.Text = "0"
88:                 Else
                        If IsNumeric(presc4.Text) Then
89:                         Me.aviam1.Focus()
                            'Me.presc5.Text = "0"
                            'Me.presc6.Text = "0"
                        Else
                            presc4.Text = ""
                        End If
                    End If

91:             Case "aviam2"
92:                 If aviam2.Text = "" Then
93:                     Me.aviam2.Text = "0"
94:                     Me.aviam3.Text = "0"
95:                     Me.aviam4.Text = "0"
a95:
96:                     Comparar()
a96:
97:                 Else
                        If IsNumeric(aviam2.Text) Then
                            If aviam2.Text >= 1111111 And aviam2.Text <= 9999999 Then
98:                             Me.aviam3.Focus()
                            End If
                        Else
                            aviam2.Text = ""
                        End If
99:                 End If
100:            Case "aviam3"
101:                If aviam3.Text = "" Then
102:                    Me.aviam3.Text = "0"
103:                    Me.aviam4.Text = "0"

104:                    Comparar()
a104:
105:                Else
                        If IsNumeric(aviam3.Text) Then
                            If aviam3.Text >= 1111111 And aviam3.Text <= 9999999 Then
106:                            Me.aviam4.Focus()
                            End If
                        Else
                            aviam3.Text = ""
                        End If
107:                End If
a107:           Case "aviam4"

z92:                If aviam4.Text = "" Then

z93:                    Me.aviam4.Text = "0"
z94:                    Comparar()

                        Me.presc1.Focus()
z96:                Else

b110:                   If IsNumeric(aviam4.Text) Then
                            If aviam4.Text >= 1111111 And aviam4.Text <= 9999999 Then
c110:
                                Comparar()
d110:
e110:                           Me.presc1.Focus()
f110:                       Else
g110:                           'Beep()
h110:                           Me.aviam4.Text = ""
i110:                           Me.aviam4.Focus()
z99:                        End If
                        Else
                            aviam4.Text = ""
                            'Comparar()

                            Me.aviam4.Focus()
                        End If
                    End If


113:                End Select
        End If
    End Sub

    'inicializa o timer para poder ter actualizações no relógio a cada segundo
    Private Sub InitializeTimer()
        On Error GoTo MOSTRARERRO
1:      ' Set to 1 second.
2:      Timer1.Interval = 1000
3:      ' Enable timer.
4:      Timer1.Enabled = True
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub InitializeTimer: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'coloca o dia da semana, data e hora na label escolhida e formata a sua apresentação
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        On Error GoTo MOSTRARERRO
1:      Me.hora.Text = Format$(Now, "ddd  dd-MM-yy  HH:mm:ss")
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Timer1_Tick: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'valida se só contém algarismos. se não dá beep e fica onde está. se sim verifica se tem 7 algarismos e passa à frente. o que acontece ao fazer enter
    Private Sub frmDesigner_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        On Error GoTo MOSTRARERRO
1:      foco = Me.ActiveControl.Name()
2:      Dim soalgarismos As Boolean
3:      If Not Char.IsNumber(e.KeyChar) Then
4:          soalgarismos = False
5:      Else
6:          soalgarismos = True
7:      End If


9:
        Select Case soalgarismos
            Case False
11:             'Beep()
12:
        End Select

35:     'retirei o "c" e o "C" como enter e foram substituidos por $m ou $J com full ascii - senão fazia shift, apareciam simbolos e andava para trás
        'If e.KeyChar = "c" Or e.KeyChar = "C" Then
        'My.Computer.Keyboard.SendKeys("{bs}")
        'My.Computer.Keyboard.SendKeys("{ENTER}")
        'End If

115:
116:
117:    'If Asc(e.KeyChar) = Keys.C Then
118:    '  My.Computer.Keyboard.SendKeys("{bs}")
119:    '  My.Computer.Keyboard.SendKeys("{ENTER}")
120:    ' End If

145:    'usados quando não tinha organismos, para permitir comparticipar tudo na paramiloidose com paramiloidoseRB
146:    'If Asc(e.KeyChar) = Keys.F11 Then
147:    'If organismosRB.Checked = True Then
148:    'organismosRB.Checked = False
149:    'but42.Checked = True
150:    'ElseIf but42.Checked = True Then
151:    'but42.Checked = False
152:    'organismosRB.Checked = True
153:    'End If
154:    'End If
155:    'usados quando não tinha organismos, para permitir comparticipar tudo na paramiloidose com paramiloidoseRB
156:    'If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.O Then
157:    ' If organismosRB.Checked = True Then
158:    ' organismosRB.Checked = False
159:    ' but42.Checked = True
160:    ' ElseIf but42.Checked = True Then
161:    ' but42.Checked = False
162:    ' organismosRB.Checked = True
163:    ' End If
164:    ' End If
165:    'usados quando não tinha organismos, para permitir comparticipar tudo na paramiloidose com paramiloidoseRB
166:    'If Asc(e.KeyChar) = Keys.Escape Then
167:    ' If organismosRB.Checked = True Then
168:    ' organismosRB.Checked = False
169:    ' but42.Checked = True
170:    ' ElseIf but42.Checked = True Then
171:    ' but42.Checked = False
172:    ' organismosRB.Checked = True
173:    ' End If
174:    ' End If
175:
176:    '   If Asc(e.KeyChar) = Keys.Space Then
177:    '    inicializar()
        '   limpar4()
        '  limpar41()
        ' CodeEC1.Text = ""
        'codEC2.Text = ""
178:    'Me.Focus()
179:    'inicializar()
        'Me.presc1.Text = ""
180:    'End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub frmDesigner_KeyPress: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub read_only()
        On Error GoTo MOSTRARERRO

        If resultadomostrado = True Then
            For Each t As TextBox In Me.Controls.OfType(Of TextBox)()
                t.ReadOnly = True
            Next
        Else
            For Each t As TextBox In Me.Controls.OfType(Of TextBox)()
                t.ReadOnly = False
            Next
        End If

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub read_only: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'os próximos 8 são os validadores que estão a funcionar. fazem saltar o focus quando se inserem 7 caracteres. e lançam a comparação
    Private Sub presc1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp1 As String
2:      Caractp1 = presc1.Text
3:
4:      If Len(presc1.Text) = 8 Then
5:          If Caractp1 Like "########" Then
6:              Prescrito1.codigo = presc1.Text
                op1.code = presc1.Text
                op1.porCNPEM = True
7:              'Me.presc2.Focus()
8:          End If
9:      End If
10:     If Len(presc1.Text) = 7 Then
            If Caractp1 Like "#######" Then
                Prescrito1.codigo = presc1.Text
                op1.code = presc1.Text
                op1.porCNPEM = False
                'Me.presc2.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub presc1_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp2 As String
2:      Caractp2 = presc2.Text
3:
4:      If Len(presc2.Text) = 8 Then
5:          If Caractp2 Like "########" Then
6:              Prescrito2.codigo = presc2.Text
                op2.code = presc2.Text
                op2.porCNPEM = True
7:              'Me.presc3.Focus()
8:          End If
9:      End If
10:     If Len(presc2.Text) = 7 Then
            If Caractp2 Like "#######" Then
                Prescrito2.codigo = presc2.Text
                op2.code = presc2.Text
                op2.porCNPEM = False
                'Me.presc3.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub presc2_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc3.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp3 As String
2:      Caractp3 = presc3.Text
3:
4:      If Len(presc3.Text) = 8 Then
5:          If Caractp3 Like "########" Then
6:              Prescrito3.codigo = presc3.Text
                op3.code = presc3.Text
                op3.porCNPEM = True
7:              'Me.presc4.Focus()
8:          End If
9:      End If
10:     If Len(presc3.Text) = 7 Then
            If Caractp3 Like "#######" Then
                : Prescrito3.codigo = presc3.Text
                op3.code = presc3.Text
                op3.porCNPEM = False
                'Me.presc4.Focus()
                : End If
            : End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub presc3_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc4.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp4 As String
2:      Caractp4 = presc4.Text
3:
4:      If Len(presc4.Text) = 8 Then
5:          If Caractp4 Like "########" Then
6:              Prescrito4.codigo = presc4.Text
                op4.code = presc4.Text
                op4.porCNPEM = True
7:              'Me.aviam1.Focus()
8:          End If
9:      End If
10:     If Len(presc4.Text) = 7 Then
            If Caractp4 Like "#######" Then
                Prescrito4.codigo = presc4.Text
                op4.code = presc4.Text
                op4.porCNPEM = False
                'Me.aviam1.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub presc4_TextChanged(: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub aviam1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta1 As String
2:      Caracta1 = aviam1.Text
3:      If Len(aviam1.Text) = 7 Then
4:          If Caracta1 Like "#######" Then
5:              Aviado1.codigo = aviam1.Text
                oa1.code = aviam1.Text
                av1.mostrado = "True"
                A = 1
                a1row = DS.infarmed.FindBycode(Aviado1.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado1.codigo)
                a1array.Add(a1row)

7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviam1_TextChanged. Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta2 As String
2:      Caracta2 = aviam2.Text
3:      If Len(aviam2.Text) = 7 Then
4:          If Caracta2 Like "#######" Then
5:              Aviado2.codigo = aviam2.Text
                oa2.code = aviam2.Text
                av2.mostrado = "True"
                A = 2
                a2row = DS.infarmed.FindBycode(Aviado2.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado2.codigo)
                a2array.Add(a2row)

7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviam2_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam3.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta3 As String
2:      Caracta3 = aviam3.Text
3:      If Len(aviam3.Text) = 7 Then
4:          If Caracta3 Like "#######" Then
5:              Aviado3.codigo = aviam3.Text
                oa3.code = aviam3.Text
                av3.mostrado = "True"
                A = 3
                a3row = DS.infarmed.FindBycode(Aviado3.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado3.codigo)
                a3array.Add(a3row)

7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviam3_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam4.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta4 As String
2:      Caracta4 = aviam4.Text
3:      If Len(aviam4.Text) = 7 Then
4:          If Caracta4 Like "#######" Then
5:              Aviado4.codigo = aviam4.Text
                oa4.code = aviam4.Text
                av4.mostrado = "True"
                A = 4
                a4row = DS.infarmed.FindBycode(Aviado4.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado4.codigo)
                a4array.Add(a4row)

8:          End If
9:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviam4_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub limparzeros()
        On Error GoTo MOSTRARERRO
1:      If presc2.Text = "0" Then
2:          presc2.BackColor = Color.White
3:      End If
4:      If presc3.Text = "0" Then
5:          presc3.BackColor = Color.White
6:      End If
7:      If presc4.Text = "0" Then
8:          presc4.BackColor = Color.White
9:      End If
10:     If aviam2.Text = "0" Then
11:         aviam2.BackColor = Color.White
12:     End If
13:     If aviam3.Text = "0" Then
14:         aviam3.BackColor = Color.White
15:     End If
16:     If aviam4.Text = "0" Then
17:         aviam4.BackColor = Color.White
18:     End If
19:     '   If presc5.Text = "0" Then
20:     '     presc5.BackColor = Color.White
21:     '    End If
22:     'If presc6.Text = "0" Then
23:     '     presc6.BackColor = Color.White
24:     '    End If
25:     '   If aviam5.Text = "0" Then
26:     '     aviam5.BackColor = Color.White
27:     '    End If
28:     '   If aviam6.Text = "0" Then
29:     '     aviam6.BackColor = Color.White
30:     '    End If

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limparzeros: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Public Sub atribuirp() 'não está a ser chamado - não serve para nada
        On Error GoTo MOSTRARERRO
1:      If P >= 4 Then
            Prescrito4.principio = p4row(1)
3:          Prescrito4.apresentacao = p4row(3)
4:          Prescrito4.dosagem = p4row(4)
5:          Prescrito4.quantidade = p4row(5)
6:          Prescrito4.comparticipacao = p4row(6)
7:          Prescrito4.grupo = p4row(7)
8:          Prescrito4.generico = p4row(10)
9:          Prescrito4.laboratorio = p4row(11)
10:     ElseIf P >= 3 Then
            Prescrito4 = vazio
12:         Prescrito3.principio = p3row(1)
13:         Prescrito3.apresentacao = p3row(3)
14:         Prescrito3.dosagem = p3row(4)
15:         Prescrito3.quantidade = p3row(5)
16:         Prescrito3.comparticipacao = p3row(6)
17:         Prescrito3.grupo = p3row(7)
18:         Prescrito3.generico = p3row(10)
19:         Prescrito3.laboratorio = p3row(11)
20:
21:     ElseIf P >= 2 Then
22:         Prescrito4 = vazio
23:         Prescrito3 = vazio
            Prescrito2.principio = p2row(1)
25:         Prescrito2.apresentacao = p2row(3)
26:         Prescrito2.dosagem = p2row(4)
27:         Prescrito2.quantidade = p2row(5)
28:         Prescrito2.comparticipacao = p2row(6)
29:         Prescrito2.grupo = p2row(7)
30:         Prescrito2.generico = p2row(10)
31:         Prescrito2.laboratorio = p2row(11)
32:
33:     Else
34:         Prescrito4 = vazio
35:         Prescrito3 = vazio
36:         Prescrito2 = vazio
            Prescrito1.principio = p1row(1)
38:         Prescrito1.apresentacao = p1row(3)
39:         Prescrito1.dosagem = p1row(4)
40:         Prescrito1.quantidade = p1row(5)
41:         Prescrito1.comparticipacao = p1row(6)
42:         Prescrito1.grupo = p1row(7)
43:         Prescrito1.generico = p1row(10)
44:         Prescrito1.laboratorio = p1row(11)
45:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub atribuirp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Public Sub atribuira() 'não está a ser chamado - não serve para nada
        On Error GoTo MOSTRARERRO
1:      If A >= 4 Then
2:          Aviado4.principio = a4row(1)
3:          Aviado4.apresentacao = a4row(3)
4:          Aviado4.dosagem = a4row(4)
5:          Aviado4.quantidade = a4row(5)
6:          Aviado4.comparticipacao = a4row(6)
7:          Aviado4.grupo = a4row(7)
8:          Aviado4.generico = a4row(10)
9:          Aviado4.laboratorio = a4row(11)
10:     ElseIf A >= 3 Then
11:         Aviado4 = vazio
            Aviado3.principio = a3row(1)
13:         Aviado3.apresentacao = a3row(3)
14:         Aviado3.dosagem = a3row(4)
15:         Aviado3.quantidade = a3row(5)
16:         Aviado3.comparticipacao = a3row(6)
17:         Aviado3.grupo = a3row(7)
18:         Aviado3.generico = a3row(10)
19:         Aviado3.laboratorio = a3row(11)
20:     ElseIf A >= 2 Then
21:         Aviado4 = vazio
22:         Aviado3 = vazio
            Aviado2.principio = a2row(1)
24:         Aviado2.apresentacao = a2row(3)
25:         Aviado2.dosagem = a2row(4)
26:         Aviado2.quantidade = a2row(5)
27:         Aviado2.comparticipacao = a2row(6)
28:         Aviado2.grupo = a2row(7)
29:         Aviado2.generico = a2row(10)
30:         Aviado2.laboratorio = a2row(11)
31:     Else
32:         Aviado4 = vazio
33:         Aviado3 = vazio
34:         Aviado2 = vazio
            Aviado1.principio = a1row(1)
36:         Aviado1.apresentacao = a1row(3)
37:         Aviado1.dosagem = a1row(4)
38:         Aviado1.quantidade = a1row(5)
39:         Aviado1.comparticipacao = a1row(6)
40:         Aviado1.grupo = a1row(7)
41:         Aviado1.generico = a1row(10)
42:         Aviado1.laboratorio = a1row(11)
43:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub atribuira: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Function diabetes(ByVal suspeito As Integer) As Boolean
1:      On Error GoTo MOSTRARERRO
2:
3:      If suspeito > 6190000 And suspeito < 6799999 Then
4:          diabetes = True
5:      Else
6:          diabetes = False
7:      End If
8:      Exit Function
MOSTRARERRO:
        MsgBox("SUB diabetes: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    'a troca de apresentação é avaliada em função da de administração (acrescentei depois o subvia que avalia a forma mesmo).
    'cada uma tem muitas formas. aqui se associa a via à forma
    Public Function via(ByVal Forma As String) As Short 'era subvia - foi acrescentado no v4.0 para substituir o via do v3.0
        On Error GoTo MOSTRARERRO
        'códigos baseados no anexo II do despacho 4586B/2013 de 1/4/2013
        'esses códigos de forma farmacêutica começam com um A seguido dos três algarismos que utilizo
        'as formas do protocólo de diabetes são as que já cá estavam
        Dim LLforma As String
        LLforma = LCase(Forma)
1:      Select Case LLforma
            Case "cápsula dura"
                via = 101
            Case "cápsula mole"
                via = 101
            Case "comprimido"
                via = 101
            Case "comprimido revestido"
                via = 101
            Case "comprimido revestido + comprimido"
                via = 101
            Case "comprimido revestido por película"
                via = 101
            Case "cápsula"
                via = 101
            Case "cápsula de libertação modificada"
                via = 102
            Case "cápsula de libertação prolongada"
                via = 102
            Case "cápsula dura de libertação prolongada"
                via = 102
            Case "cápsula mole de libertação modificada"
                via = 102
            Case "comprimido de libertação modificada"
                via = 102
            Case "comprimido de libertação prolongada"
                via = 102
            Case "comprimido de libertação prolongada revestido por película"
                via = 102
            Case "cápsula dura de libertação modificada"
                via = 102
            Case "cápsula gastrorresistente"
                via = 103
            Case "cápsula mole gastrorresistente"
                via = 103
            Case "comprimido gastrorresistente"
                via = 103
            Case "cápsula dura gastrorresistente"
                via = 103
            Case "comprimido para chupar"
                via = 106
            Case "comprimido para mastigar"
                via = 106
            Case "goma para mascar medicamentosa"
                via = 106
            Case "liofilizado oral"
                via = 106
            Case "película orodispersível"
                via = 106
            Case "comprimido orodispersível"
                via = 106
            Case "pastilha"
                via = 108
            Case "comprimido sublingual"
                via = 109
            Case "comprimido + suspensão oral"
                via = 110
            Case "gel oral"
                via = 111
            Case "comprimido dispersível"
                via = 112
            Case "comprimido efervescente"
                via = 112
            Case "comprimido solúvel"
                via = 112
            Case "gotas orais, suspensão"
                via = 112
            Case "granulado"
                via = 112
            Case "granulado efervescente"
                via = 112
            Case "granulado para solução oral"
                via = 112
            Case "granulado para suspensão oral"
                via = 112
            Case "granulado para xarope"
                via = 112
            Case "pó e solvente para suspensão oral"
                via = 112
            Case "pó efervescente"
                via = 112
            Case "pó oral"
                via = 112
            Case "pó para solução oral"
                via = 112
            Case "pó para suspensão oral"
                via = 112
            Case "solução oral"
                via = 112
            Case "xarope"
                via = 112
            Case "gotas orais, solução"
                via = 112
            Case "pó e solvente para solução oral"
                via = 112
            Case "suspensão oral"
                via = 112

            Case "solução para pulverização bucal"
                via = 115
            Case "suspensão dental"
                via = 116
            Case "granulado de libertação modificada"
                via = 117
            Case "granulado de libertação prolongada"
                via = 117
            Case "pasta dentífrica"
                via = 118
            Case "gel dental"
                via = 119
            Case "champô"
                via = 200
            Case "espuma cutânea"
                via = 201
            Case "líquido cutâneo"
                via = 201
            Case "solução cutânea"
                via = 201
            Case "emulsão cutânea"
                via = 201
            Case "gel"
                via = 202
            Case "pomada"
                via = 202
            Case "creme"
                via = 202
            Case "pó para pulverização cutânea"
                via = 203
            Case "solução para pulverização cutânea"
                via = 203
            Case "pó cutâneo"
                via = 203
            Case "sistema transdérmico"
                via = 204
            Case "penso impregnado"
                via = 204
            Case "colírio, solução"
                via = 300
            Case "colírio, suspensão"
                via = 300
            Case "colírio, pó e solvente para solução"
                via = 300
            Case "colírio de libertação prolongada"
                via = 301
            Case "colírio de acção prolongada"
                via = 301
            Case "pomada oftálmica"
                via = 302
            Case "gel oftálmico"
                via = 302
            Case "gotas auriculares, suspensão"
                via = 400
            Case "solução para pulverização auricular"
                via = 400
            Case "gotas auriculares, solução"
                via = 400
            Case "gás para inalação"
                via = 500

            Case "solução para pulverização nasal"
                via = 501
            Case "suspensão para pulverização nasal"
                via = 501
            Case "gotas nasais, solução"
                via = 501

                'infarmed descodificou do 503
            Case "cápsula para inalação"
                via = 504
            Case "pó nasal"
                via = 503
            Case "pó para inalação"
                via = 503
            Case "cápsula para inalação por vaporização"
                via = 503
            Case "pó para inalação, cápsula"
                via = 503
            Case "pó para inalação, cápsula dura"
                via = 503
            Case "solução para inalação por nebulização"
                via = 503
            Case "solução para inalação por vaporização"
                via = 503
            Case "solução pressurizada para inalação"
                via = 503
            Case "suspensão" 'descontinuado a 201309 maas deixei ficar na mesma com 503
                via = 503
            Case "suspensão pressurizada para inalação"
                via = 503
            Case "líquido para inalação por vaporização"
                via = 503
            Case "suspensão para inalação por nebulização"
                via = 503

                ' infarmed não codificou por isso criei eu - não pus no 503 para poder dar erro
            Case "pó para inalação em recipiente unidose"
                via = 504

            Case "comprimido vaginal"
                via = 600
            Case "cápsula mole vaginal"
                via = 600
            Case "creme vaginal"
                via = 602
            Case "espuma vaginal"
                via = 602
            Case "gel vaginal"
                via = 602
            Case "óvulo"
                via = 602
            Case "pomada vaginal"
                via = 602
            Case "solução vaginal"
                via = 602
            Case "creme vaginal + óvulo"
                via = 602
            Case "dispositivo intra-uterino"
                via = 603

                'infarmed descodificou do 603
            Case "dispositivo de libertação intra-uterino"
                via = 604
            Case "comprimido para suspensão rectal"
                via = 700
            Case "enema, solução"
                via = 700
            Case "espuma rectal"
                via = 700
            Case "pomada rectal"
                via = 700
            Case "pomada rectal + supositório"
                via = 700
            Case "solução rectal"
                via = 700
            Case "suspensão rectal"
                via = 700
            Case "enema, suspensão"
                via = 700
            Case "supositório"
                via = 704
            Case "cápsula dura + pó e solvente para solução injectável"
                via = 800
            Case "concentrado e solvente para solução para perfusão"
                via = 800
            Case "concentrado para solução injectável"
                via = 800
            Case "concentrado para solução injectável ou para perfusão"
                via = 800
            Case "concentrado para solução para perfusão"
                via = 800
            Case "emulsão injectável"
                via = 800
            Case "emulsão para perfusão"
                via = 800
            Case "liofilizado para solução para perfusão"
                via = 800
            Case "pó e solvente para solução injectável"
                via = 800
            Case "pó e solvente para solução injectável ou para perfusão"
                via = 800
            Case "pó e solvente para solução para perfusão"
                via = 800
            Case "pó e solvente para suspensão injectável"
                via = 800
            Case "pó e veículo para suspensão injectável"
                via = 800
            Case "pó para concentrado para solução injectável ou para perfusão"
                via = 800
            Case "pó para concentrado para solução para perfusão"
                via = 800
            Case "pó para solução injectável"
                via = 800
            Case "pó para solução injectável ou para perfusão"
                via = 800
            Case "pó para solução ou para suspensão injectável"
                via = 800
            Case "pó para solução para perfusão"
                via = 800
            Case "solução injectável"
                via = 800
            Case "solução injectável ou concentrado para solução para perfusão"
                via = 800

            Case "solvente/veículo para uso parentérico"
                via = 800
            Case "suspensão injectável"
                via = 800

            Case "emulsão injectável ou para perfusão"
                via = 800

            Case "solução injectável ou para perfusão"
                via = 800
            Case "solução para perfusão"
                via = 800

            Case "pó para suspensão para implantação"
                via = 803
            Case "solução para diálise peritoneal"
                via = 806
            Case "comprimido + supositório"
                via = 900
            Case "implante"
                via = 901
            Case "verniz medicamentoso para as unhas"
                via = 902
            Case "verniz para as unhas medicamentoso"
                via = 902
272:        Case "seringa"
273:
                via = 21
274:        Case "agulhas"
275:
                via = 22
276:        Case "lancetas"

277:            via = 23
278:        Case "lancetas para punção capilar"

                via = 23
279:        Case "lancetas para amostragem de sangue"

                via = 23
280:        Case "lancetas esterilizadas por irradiação, de uso único"

281:            via = 23
283:        Case "lancetas estéreis"

284:            via = 23
285:        Case "lancetas estéreis para obtenção de uma gota de sangue"

286:            via = 23
287:        Case "lancetas estéreis por radiação gama"

288:            via = 23
289:        Case "lancetas esterilizadas por radiação gama"

290:            via = 23
291:        Case "seringa de uso único, estéril para administração de insulina com agulha ultrafina, 30G, 8mm"

292:            via = 21
293:        Case "seringas de 0,3 ml (8mm) diametro 0,3mm (30G) escala de 30 unidades divididas em 1/2"

294:            via = 21
295:        Case "tiras para determinação de glicémia"

296:            via = 24
297:        Case "tiras para determinação de glicosúria"

298:            via = 25
299:        Case "tiras para determinação de glicosúria e cetonúria"

300:            via = 26
301:        Case "tiras teste de b-cetonemia"

302:            via = 27
303:        Case "agulha de 0,30x8mm (30 Gx5/16)"

304:            via = 22
305:        Case "agulha de 0,33x12mm (29 Gx15/32)"

306:            via = 22
307:        Case "agulha de uso único, estéril, para canetas de administração de insulina"

308:            via = 22
309:
509:        Case Else
510:            'MsgBox("forma farmacêutica desconhecida")
511:            End Select
512:    If via = 0 Then
513:        ' MsgBox("via = 0")
514:    End If
515:    Exit Function
MOSTRARERRO:
        MsgBox("function via: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    'chama o irbuscar() e o row2array para produzir os resultados
    Private Sub Comparar()
        On Error GoTo MOSTRARERRO
        If comparado = False Then
1:          Me.result1.Text = ""
2:          Me.result1.BackColor = cordefundo
3:          Me.result2.Text = ""
4:          Me.result2.BackColor = cordefundo
5:          Me.result3.Text = ""
6:          Me.result3.BackColor = cordefundo
7:          Me.result4.Text = ""
8:          Me.result4.BackColor = cordefundo
9:          irbuscar2013()
10:         'row2array()
11:         comparador2013() '2013
            makeranking2013()
            'prioridade2013()
12:
            novoprioridade2013()

            mostrarcodigo()
1901:
1905:       limpo = False

20:         comparado = True
            inicializado = False
            read_only()
        End If
SAIR:
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Comparar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub novoprioridade2013()
1:      On Error GoTo MOSTRARERRO
2:
3:      Dim iterPPP As Single = 1
4:      Dim iterAAA As Single = 1
5:      Dim arraycroces(45) As Object
6:      Dim concatP1 As String
7:      Dim concatP2 As String
8:      Dim concatP3 As String
9:      Dim concatP4 As String
10:     Dim concatA1 As String
11:     Dim concatA2 As String
12:     Dim concatA3 As String
13:     Dim concatA4 As String
14:     concatP1 = 0
15:     concatP2 = 0
16:     concatP3 = 0
17:     concatP4 = 0
18:     concatA1 = 0
19:     concatA2 = 0
20:     concatA3 = 0
21:     concatA4 = 0
26:     arraycroces(11) = cruzamento11
27:     arraycroces(12) = cruzamento12
28:     arraycroces(13) = cruzamento13
29:     arraycroces(14) = cruzamento14
30:     arraycroces(21) = cruzamento21
31:     arraycroces(22) = cruzamento22
32:     arraycroces(23) = cruzamento23
33:     arraycroces(24) = cruzamento24
34:     arraycroces(31) = cruzamento31
35:     arraycroces(32) = cruzamento32
36:     arraycroces(33) = cruzamento33
37:     arraycroces(34) = cruzamento34
38:     arraycroces(41) = cruzamento41
39:     arraycroces(42) = cruzamento42
40:     arraycroces(43) = cruzamento43
41:     arraycroces(44) = cruzamento44
42:
46:     Dim cruzamentocerto(3) As Object
47:     cruzamentocerto(0) = resultado1
48:     cruzamentocerto(1) = resultado2
49:     cruzamentocerto(2) = resultado3
50:     cruzamentocerto(3) = resultado4

51:     Dim iteradorPrior As Single = 11
52:     Dim iteradorQual1 As Single = 11
53:     Dim iteradorQual2 As Single = 11
54:
55:     Dim arraydciP(3) As String
56:     Dim arraydciA(3) As String
57:
58:
70:
72:     Dim arrayoa(3) As Object
73:     arrayoa(0) = oa1
74:     arrayoa(1) = oa2
75:     arrayoa(2) = oa3
76:     arrayoa(3) = oa4
77:     Dim arrayop(3) As Object
78:     arrayop(0) = op1
79:     arrayop(1) = op2
80:     arrayop(2) = op3
81:     arrayop(3) = op4
82:     Dim quantosdciAnosP(3) As Single
83:     quantosdciAnosP(0) = quantosDCIa1nosP
84:     quantosdciAnosP(1) = quantosDCIa2nosP
85:     quantosdciAnosP(2) = quantosDCIa3nosP
86:     quantosdciAnosP(3) = quantosDCIa4nosP
87:     Dim quantosdciAnosA(3) As Single
88:     quantosdciAnosA(0) = quantosDCIa1
89:     quantosdciAnosA(1) = quantosDCIa2
90:     quantosdciAnosA(2) = quantosDCIa3
91:     quantosdciAnosA(3) = quantosDCIa4
92:
93:     Dim concatP(3) As String
94:     concatP(0) = concatP1
95:     concatP(1) = concatP2
96:     concatP(2) = concatP3
97:     concatP(3) = concatP4
98:     Dim concatA(3) As String
99:     concatA(0) = concatA1
100:    concatA(1) = concatA2
101:    concatA(2) = concatA3
102:    concatA(3) = concatA4
103:
104:    For n As Integer = 1 To P
105:        arraydciP(n - 1) = arrayop(n - 1).dci
106:    Next
107:    For nn As Integer = 1 To A
108:        arraydciA(nn - 1) = arrayoa(nn - 1).dci
109:    Next

        For nnn As Integer = 1 To A
            For nnnn As Integer = 1 To P
110:            If arraydciP(nnnn - 1) = arrayoa(nnn - 1).dci Then
                    quantosdciAnosP(nnn - 1) = quantosdciAnosP(nnn - 1) + 1
                End If
            Next
            For aaaa As Integer = 1 To A
                If arraydciA(aaaa - 1) = arrayoa(nnn - 1).dci Then
                    quantosdciAnosA(nnn - 1) = quantosdciAnosA(nnn - 1) + 1
                End If
            Next
            concatA(nnn - 1) = arrayoa(nnn - 1).dci & arrayoa(nnn - 1).dose & arrayoa(nnn - 1).forma & arrayoa(nnn - 1).qty
        Next

        For nnnnn As Integer = 1 To P
            concatP(nnnnn - 1) = arrayop(nnnnn - 1).dci & arrayop(nnnnn - 1).dose & arrayop(nnnnn - 1).forma & arrayop(nnnnn - 1).qty
        Next



124:    Do While iterAAA <= A
125:
126:        Select Case quantosdciAnosP(iterAAA - 1)
                Case Is = 0
128:                cruzamentocerto(iterAAA - 1).ranking = "000000" 'AVIADO NÃO PRESCRiTO
129:                'MsgBox("0 dci A nos P")
134:            Case Else
135:                SUBROTINA(iterPPP, iterAAA, iteradorPrior, iteradorQual1, iteradorQual2, arraycroces, cruzamentocerto, quantosdciAnosP(iterAAA - 1), quantosdciAnosA(iterAAA - 1))
136:
            End Select
137:
146:
            If aceitarduplicados = True Then
150:            'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
165:            For AAA As Integer = 1 To A
169:                If AAA <> iterAAA Then
                        If (arrayoa(AAA - 1).cnpem > 0 And (arrayoa(AAA - 1).cnpem = arrayoa(iterAAA - 1).cnpem)) Or (arrayoa(AAA - 1).code = arrayoa(iterAAA - 1).code) Then
170:                        If iterAAA > AAA Then 'para não fazer duas vezes e desfazer o que estava feito
                                cruzamentocerto(iterAAA - 1).ranking = cruzamentocerto(AAA - 1).ranking
                            End If
172:                    End If
173:                End If
174:            Next
            End If
152:
159:        iterAAA = iterAAA + 1
160:    Loop 'final do while iterAAA
161:
162:
163:
164:
166:
176:
177:
178:    Exit Sub
MOSTRARERRO:
        MsgBox("sub novoprioridade2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Dim ordenado1(3)
    Dim ordenado2(3)
    Dim ordenado3(3)
    Dim ordenado4(3)

    Sub SUBROTINA(ByVal ITERppp As Single, ByVal ITERaaa As Single, ByVal ITERadorprior As Single, ByVal iteradorQual1 As Single, ByVal iteradorQual2 As Single, ByVal arraycroces As Array, ByVal cruzamentocerto As Array, ByVal quantosdciAnosP As Single, ByVal quantosdciAnosA As Single)
1:
        On Error GoTo MOSTRARERRO
        cruzamentocerto(0) = resultado1
        cruzamentocerto(1) = resultado2
        cruzamentocerto(2) = resultado3
        cruzamentocerto(3) = resultado4
2:      Dim ordenados(3) As Array
3:      ordenados(0) = ordenado1
4:      ordenados(1) = ordenado2
5:      ordenados(2) = ordenado3
6:      ordenados(3) = ordenado4
7:
92:     'If arraycroces(ITERadorprior).anulado = False Then
93:     ITERppp = 1
94:     Do While ITERppp <= P
361:
363:        ITERadorprior = Convert.ToSingle((ITERppp) & (ITERaaa))
37:         If arraycroces(ITERadorprior).anulado = False Then
38:             If arraycroces(ITERadorprior).ranking.ToString.Substring(1) >= cruzamentocerto(ITERaaa - 1).ranking.ToString.Substring(1) Then
39:                 'só entra uma vez
                    cruzamentocerto(ITERaaa - 1).qualP = ITERppp
391:                cruzamentocerto(ITERaaa - 1).ranking = arraycroces(ITERadorprior).ranking
393:                iteradorQual1 = Convert.ToSingle(cruzamentocerto(ITERaaa - 1).qualP & ITERaaa)
394:                iteradorQual2 = iteradorQual1
42:             End If
43:
45:         End If
46:
47:
48:         ITERppp = ITERppp + 1
49:     Loop 'final do while iterPPP
50:     anular(iteradorQual1, iteradorQual2, ITERaaa, arraycroces, cruzamentocerto)
51:
54:     Exit Sub
MOSTRARERRO:
        MsgBox("sub subrotina: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Sub anular(ByVal iteradorqual1 As Single, iteradorqual2 As Single, ByVal iterAAA As Single, ByVal arraycroces As Object, ByVal cruzamentocerto As Object)
        On Error GoTo MOSTRARERRO

6:      arraycroces(11) = cruzamento11
7:      arraycroces(12) = cruzamento12
8:      arraycroces(13) = cruzamento13
9:      arraycroces(14) = cruzamento14
10:     arraycroces(21) = cruzamento21
11:     arraycroces(22) = cruzamento22
12:     arraycroces(23) = cruzamento23
13:     arraycroces(24) = cruzamento24
14:     arraycroces(31) = cruzamento31
15:     arraycroces(32) = cruzamento32
16:     arraycroces(33) = cruzamento33
17:     arraycroces(34) = cruzamento34
18:     arraycroces(41) = cruzamento41
19:     arraycroces(42) = cruzamento42
20:     arraycroces(43) = cruzamento43
21:     arraycroces(44) = cruzamento44

50:     Do While iteradorqual1 <= P * 10 + iterAAA
51:         arraycroces(iteradorqual1).anulado = True
52:         iteradorqual1 = iteradorqual1 + 10
53:     Loop
54:
55:     Do While iteradorqual2 <= cruzamentocerto(iterAAA - 1).qualP * 10 + A
            If aceitarduplicados = False Then
                'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
56:             arraycroces(iteradorqual2).anulado = True 'tirar plica para implicar com duplicados - o codigo da duplicação esta no novoprioridade2013 e no mostrarcodigo
57:         End If
            iteradorqual2 = iteradorqual2 + 1
58:     Loop
        Exit Sub
MOSTRARERRO:
        MsgBox("sub anular: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub prioridade2013()
1:      On Error GoTo MOSTRARERRO
2:
3:      Dim iterPPP As Single = 1
4:      Dim iterAAA As Single = 1
        Dim arraycroces(45) As Object

6:      arraycroces(11) = cruzamento11
7:      arraycroces(12) = cruzamento12
8:      arraycroces(13) = cruzamento13
9:      arraycroces(14) = cruzamento14
10:     arraycroces(21) = cruzamento21
11:     arraycroces(22) = cruzamento22
12:     arraycroces(23) = cruzamento23
13:     arraycroces(24) = cruzamento24
14:     arraycroces(31) = cruzamento31
15:     arraycroces(32) = cruzamento32
16:     arraycroces(33) = cruzamento33
17:     arraycroces(34) = cruzamento34
18:     arraycroces(41) = cruzamento41
19:     arraycroces(42) = cruzamento42
20:     arraycroces(43) = cruzamento43
21:     arraycroces(44) = cruzamento44
22:
26:     Dim cruzamentocerto(3) As Object
27:     cruzamentocerto(0) = resultado1
28:     cruzamentocerto(1) = resultado2
29:     cruzamentocerto(2) = resultado3
30:     cruzamentocerto(3) = resultado4

31:     Dim iteradorPrior As Single = 11
        Dim iteradorQual1 As Single = 11
        Dim iteradorQual2 As Single = 11


35:     Do While iterAAA <= A
            iterPPP = 1
36:         Do While iterPPP <= P
361:            'MsgBox("P: " & iterPPP)
                'MsgBox("a: " & iterAAA)
363:            iteradorPrior = Convert.ToSingle((iterPPP) & (iterAAA))

37:             If arraycroces(iteradorPrior).anulado = False Then
38:                 If arraycroces(iteradorPrior).ranking.ToString.Substring(1) >= cruzamentocerto(iterAAA - 1).ranking.ToString.Substring(1) Then
39:                     'MsgBox(iteradorPrior & ": " & arraycroces(iteradorPrior).ranking)
                        cruzamentocerto(iterAAA - 1).qualP = iterPPP
391:                    cruzamentocerto(iterAAA - 1).ranking = arraycroces(iteradorPrior).ranking
393:                    iteradorQual1 = Convert.ToSingle(cruzamentocerto(iterAAA - 1).qualP & iterAAA)
394:                    iteradorQual2 = iteradorQual1
40:                 Else
41:                     'fica como estava
42:                 End If
43:             Else
44:                 'fica como estava
45:             End If
46:
47:
48:             iterPPP = iterPPP + 1
49:         Loop 'final do while iterPPP

50:         Do While iteradorQual1 <= P * 10 + iterAAA
51:             arraycroces(iteradorQual1).anulado = True
52:             iteradorQual1 = iteradorQual1 + 10
53:         Loop
54:
55:         Do While iteradorQual2 <= cruzamentocerto(iterAAA - 1).qualP * 10 + A
56:             arraycroces(iteradorQual2).anulado = True
57:             iteradorQual2 = iteradorQual2 + 1
58:         Loop


59:         iterAAA = iterAAA + 1
60:     Loop 'final do while iterAAA
61:
62:
63:     Exit Sub
MOSTRARERRO:
        MsgBox("prioridade2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'lite não tem o que vinha a seguir ao OK:
    Sub makeranking2013()

1:      On Error GoTo MOSTRARERRO
2:      Dim iterPP As Single = 1
3:      Dim iterAA As Single = 1
        Dim arraycruzamentos(44) As Object
5:      arraycruzamentos(11) = cruzamento11
6:      arraycruzamentos(12) = cruzamento12
7:      arraycruzamentos(13) = cruzamento13
8:      arraycruzamentos(14) = cruzamento14
9:      arraycruzamentos(21) = cruzamento21
10:     arraycruzamentos(22) = cruzamento22
11:     arraycruzamentos(23) = cruzamento23
12:     arraycruzamentos(24) = cruzamento24
13:     arraycruzamentos(31) = cruzamento31
14:     arraycruzamentos(32) = cruzamento32
15:     arraycruzamentos(33) = cruzamento33
16:     arraycruzamentos(34) = cruzamento34
17:     arraycruzamentos(41) = cruzamento41
18:     arraycruzamentos(42) = cruzamento42
19:     arraycruzamentos(43) = cruzamento43
20:     arraycruzamentos(44) = cruzamento44
        Dim arrayp(3) As Object
        arrayp(0) = op1
        arrayp(1) = op2
        arrayp(2) = op3
        arrayp(3) = op4
        Dim ARRAYdciP(3) As String 'não está no não lite?
        Dim iteradorMR As Single = 11

21:     Do While iterAA <= A
211:        iterPP = 1
22:         Do While iterPP <= P
221:            iteradorMR = Convert.ToSingle((iterPP) & (iterAA))
23:             'ranking do code c*********
24:             If arraycruzamentos(iteradorMR).code = True Then 'se por mesmo código
25:                 arraycruzamentos(iteradorMR).ranking = "411123"
26:                 GoTo OK
27:             ElseIf arrayp(iterPP - 1).porCNPEM = False Then 'se code é dif e não prescrito por cnpem
271:                'será como se xcnpem =1 ou 3
272:                If arrayp(iterPP - 1).cnpem <> 0 Then 'prescrito por code diferente e sem CNPEM
273:                    arraycruzamentos(iteradorMR).ranking = "2" 'comparar à antiga, pode estar certo ou errado
274:                Else 'prescrito por code diferente com CNPEM
275:                    If arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por code dif com mesmo cnpem (pode estar certo ou S)
276:                        If arraycruzamentos(iteradorMR).marcadifsemhavergens = True Then
277:                            arraycruzamentos(iteradorMR).ranking = "211123"
278:                            GoTo OK
279:                        Else
2791:
2792:                           If arrayp(iterPP - 1).lab = "0" Then
280:                                arraycruzamentos(iteradorMR).ranking = "2"
2801:                           Else
2802:                               arraycruzamentos(iteradorMR).ranking = "211103" 'marcadif
2803:                               GoTo OK
2804:                           End If
2805:
281:                        End If
282:                    Else
28:                         arraycruzamentos(iteradorMR).ranking = "1"  'por code dif com cnpem dif - está errado -  vai seguir o restante
29:                     End If
291:                    End If

292:            ElseIf arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por cnpem e igual 'está diferente no não lite
30:                 arraycruzamentos(iteradorMR).ranking = "311123"
31:                 GoTo OK
32:             Else 'prescrito por cnpem e diferente (xcnpem = 2 ou 0)
33:                 arraycruzamentos(iteradorMR).ranking = "0" 'prescrito por cnpem e diferente

34:             End If
35:
36:             'ranking do dci *d********
37:             If arraycruzamentos(iteradorMR).dci = True And arrayp(iterPP - 1).dci <> "0" Then
38:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
39:             Else
40:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
41:             End If
42:
43:             'ranking da forma **f*******
44:             If arraycruzamentos(iteradorMR).forma = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'diferente no não lite
                    If Not (arraycruzamentos(iteradorMR).cnpem = False And arraycruzamentos(iteradorMR).dci = True And arraycruzamentos(iteradorMR).dose = True And arraycruzamentos(iteradorMR).xqty = 3) Then 'para evitar dar certo com cnpem diferente devido a forma igual
45:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
                    Else
                        arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
                    End If
46:             Else
47:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
48:             End If
49:
50:             'ranking da dose ***d******
51:             If arraycruzamentos(iteradorMR).dose = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'diferente no não lite
52:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
53:             Else
54:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
55:             End If

56:
57:             'ranking da marca ****n*****
                arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).marcadifsemhavergens

63:
64:             'ranking da qty *****q****
65:             arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).xqty
66:


OK:             iterPP = iterPP + 1
77:         Loop
78:         iterAA = iterAA + 1
79:     Loop

80:     Exit Sub
MOSTRARERRO:
        MsgBox("makeranking2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub comparador2013() '2013
1:      On Error GoTo MOSTRARERRO
        Dim posicaop1(3) As Single
        Dim posicaop2(3) As Single
        Dim antesp(3) As Double
        Dim depoisp(3) As Double
        Dim posicaoa1(3) As Single
        Dim posicaoa2(3) As Single
        Dim antesa(3) As Double
        Dim depoisa(3) As Double
2:      Dim arrayA(3) As Object
3:      Dim arrayP(3) As Object
4:      Dim arraycruz(44) As Object '{11, 12, 13, 14, 21, 22, 23, 24, 31, 32, 33, 34, 41, 42, 43, 44}
6:      Dim iterP As Single = 1
7:      Dim iterA As Single = 1
        Dim rowP(3) As Object
        rowP(0) = p1row
        rowP(1) = p2row
        rowP(2) = p3row
        rowP(3) = p4row
        Dim rowA(3) As Object
        rowA(0) = a1row
        rowA(1) = a2row
        rowA(2) = a3row
        rowA(3) = a4row
9:      arrayA(0) = oa1
10:     arrayA(1) = oa2
11:     arrayA(2) = oa3
12:     arrayA(3) = oa4
13:     arrayP(0) = op1
14:     arrayP(1) = op2
15:     arrayP(2) = op3
16:     arrayP(3) = op4
17:     arraycruz(11) = cruzamento11
18:     arraycruz(12) = cruzamento12
19:     arraycruz(13) = cruzamento13
20:     arraycruz(14) = cruzamento14
21:     arraycruz(21) = cruzamento21
22:     arraycruz(22) = cruzamento22
23:     arraycruz(23) = cruzamento23
24:     arraycruz(24) = cruzamento24
25:     arraycruz(31) = cruzamento31
26:     arraycruz(32) = cruzamento32
27:     arraycruz(33) = cruzamento33
28:     arraycruz(34) = cruzamento34
29:     arraycruz(41) = cruzamento41
30:     arraycruz(42) = cruzamento42
31:     arraycruz(43) = cruzamento43
32:     arraycruz(44) = cruzamento44

331:    Dim iterador As Single = 11
332:    Do While iterA <= A
333:        'If Not IsNothing(rowA(iterA)) Then
334:        iterP = 1
341:        Do While iterP <= P
342:            'If Not IsNothing(rowP(iterP)) Then
343:
345:            iterador = Convert.ToSingle((iterP) & (iterA))
35:             If arrayA(iterA - 1).code = arrayP(iterP - 1).code Then
36:                 arraycruz(iterador).code = True
37:             Else
38:                 arraycruz(iterador).code = False
39:             End If

40:             If arrayA(iterA - 1).dci = arrayP(iterP - 1).dci Then
41:                 arraycruz(iterador).dci = True
42:             Else
43:                 arraycruz(iterador).dci = False
44:             End If

45:             If arrayA(iterA - 1).nome = arrayP(iterP - 1).nome Then
46:                 arraycruz(iterador).nome = True
47:             Else
48:                 arraycruz(iterador).nome = False
49:             End If

               
50:             If arrayA(iterA - 1).forma = arrayP(iterP - 1).forma Then
51:                 arraycruz(iterador).forma = True
52:             Else
53:                 arraycruz(iterador).forma = False
54:             End If
             
68:             If arraycruz(iterador).forma = False Then 'caso particular das formas que podem ser equivalentes (não implica que a regra do CNPEM não acuse erro por causa da forma, por exemplo creme e pomada)
69:                 If via(arrayP(iterP - 1).forma) = via(arrayA(iterA - 1).forma) Then
70:                     arraycruz(iterador).forma = True
71:                 End If
72:             End If

721:            If arrayA(iterA - 1).dose = arrayP(iterP - 1).dose Then
722:                arraycruz(iterador).dose = True
723:            Else
724:                arraycruz(iterador).dose = False
725:            End If
                'até linha 865 não estáVA no 6.0 lite (GH, gen ,lab, dci_obr, qty)





726:            If arrayA(iterA - 1).qty = arrayP(iterP - 1).qty Then
727:                arraycruz(iterador).qty = True
728:            Else
729:                arraycruz(iterador).qty = False
730:            End If

73:             If IsNumeric(arrayP(iterP - 1).qty) AndAlso IsNumeric(arrayA(iterA - 1).qty) Then 'qty numerica em ambos
731:                If arraycruz(iterador).qty = False Then 'se a quantidade não é igual vai ver se é menor, maior e maior que 50% (independentemente CNPEM)
74:                     If Convert.ToInt16(arrayP(iterP - 1).qty) > Convert.ToInt16(arrayA(iterA - 1).qty) Then 'se qty aviado < prescrito
75:                         arraycruz(iterador).xqty = 2
76:                     ElseIf Convert.ToInt16(arrayP(iterP - 1).qty) > 1.5 * Convert.ToInt16(arrayA(iterA - 1).qty) Then 'se qty aviado > 150% prescrito
77:                         arraycruz(iterador).xqty = 0
78:                     ElseIf Convert.ToInt16(arrayP(iterP - 1).qty) < 1.5 * Convert.ToInt16(arrayA(iterA - 1).qty) And Convert.ToInt16(arrayA(iterA - 1).qty) > Convert.ToInt16(arrayP(iterP - 1).qty) Then 'se qty prescrito < qty aviado < 150 % prescrito
79:                         arraycruz(iterador).xqty = 1
80:                     End If
81:                 Else
82:                     arraycruz(iterador).xqty = 3 'qty =     (independentemente CNPEM) 
83:                 End If

801:            Else 'qty não numerica em pelo menos um dos dois
                    If IsNumeric(arrayP(iterP - 1).qty) Or IsNumeric(arrayA(iterA - 1).qty) Then 'qty numerica num deles
                        'é erro de apresentacao.
8147:                   arraycruz(iterador).xqty = 0
                    Else 'nenhum é numerico
                        If Len(arrayP(iterP - 1).qty) >= 17 Then
802:                        posicaop1(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("-")
803:                        If arrayP(iterP - 1).qty.ToString.Contains("o") Then
804:                            posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("o")
805:                        ElseIf arrayP(iterP - 1).qty.ToString.Contains("m") Then
806:                            posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("m")
807:                        ElseIf arrayP(iterP - 1).qty.ToString.Contains("g") Then
808:                            posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("g")
809:                        ElseIf arrayP(iterP - 1).qty.ToString.Contains("l") Then
810:                            posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("l")
811:                        End If
812:                        antesp(iterP - 1) = arrayP(iterP - 1).qty.ToString.Substring(0, posicaop1(iterP - 1) - 12)
813:                        depoisp(iterP - 1) = arrayP(iterP - 1).qty.ToString.Substring(posicaop1(iterP - 1) + 2, posicaop2(iterP - 1) - posicaop1(iterP - 1) - 3)
814:                       
                        End If
                        If Len(arrayA(iterA - 1).qty) >= 17 Then
815:                        posicaoa1(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("-")
816:                        If arrayA(iterA - 1).qty.ToString.Contains("o") Then
817:                            posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("o")
818:                        ElseIf arrayA(iterA - 1).qty.ToString.Contains("m") Then
819:                            posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("m")
820:                        ElseIf arrayA(iterA - 1).qty.ToString.Contains("g") Then
821:                            posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("g")
822:                        ElseIf arrayA(iterA - 1).qty.ToString.Contains("l") Then
823:                            posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("l")
824:                        End If

825:                        antesa(iterA - 1) = arrayA(iterA - 1).qty.ToString.Substring(0, posicaoa1(iterA - 1) - 12)
826:                        depoisa(iterA - 1) = arrayA(iterA - 1).qty.ToString.Substring(posicaoa1(iterA - 1) + 2, posicaoa2(iterA - 1) - posicaoa1(iterA - 1) - 3)
827:
                        End If
828:                    'compara antesp(iterP - 1) com antesa(iterA - 1)
829:                    If antesp(iterP - 1) <> antesa(iterA - 1) Then 'se a quantidade não numerica não é igual vai ver se é menor, maior e maior que 50% (independentemente CNPEM)
830:                        If Convert.ToInt16(antesp(iterP - 1)) > Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado < prescrito
831:                            arraycruz(iterador).xqty = 2
832:                        ElseIf Convert.ToInt16(antesp(iterP - 1)) > 1.5 * Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado > 150% prescrito
833:                            arraycruz(iterador).xqty = 0
834:                        ElseIf Convert.ToInt16(antesp(iterP - 1)) < 1.5 * Convert.ToInt16(antesa(iterA - 1)) And Convert.ToInt16(antesa(iterA - 1)) > Convert.ToInt16(antesp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
835:                            arraycruz(iterador).xqty = 1
8136:                       End If
8137:                   Else 'se antes é igual compara depoisp(iterP - 1) com depoisa(iterA - 1)
8138:                       If Convert.ToInt16(depoisp(iterP - 1)) > Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado < prescrito
8139:                           arraycruz(iterador).xqty = 2
8140:                       ElseIf Convert.ToInt16(depoisp(iterP - 1)) > 1.5 * Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado > 150% prescrito
8141:                           arraycruz(iterador).xqty = 0
8142:                       ElseIf Convert.ToInt16(depoisp(iterP - 1)) < 1.5 * Convert.ToInt16(depoisa(iterA - 1)) And Convert.ToInt16(depoisa(iterA - 1)) > Convert.ToInt16(depoisp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
8143:                           arraycruz(iterador).xqty = 1
                            Else
81433:                          'tanto o antes foi igual como o depois também agora é igual
81434:                          arraycruz(iterador).xqty = 3
8144:                       End If
8145:                   End If
8146:               End If
8148:           End If



836:            If arrayA(iterA - 1).GH = arrayP(iterP - 1).GH Then
837:                arraycruz(iterador).GH = True
838:            Else
839:                arraycruz(iterador).GH = False
840:            End If


846:            If arrayA(iterA - 1).gen = arrayP(iterP - 1).gen Then
847:                arraycruz(iterador).gen = True
848:            Else
849:                arraycruz(iterador).gen = False
850:            End If

851:            If arrayA(iterA - 1).lab = arrayP(iterP - 1).lab Then
852:                arraycruz(iterador).lab = True
853:            Else
854:                arraycruz(iterador).lab = False
855:            End If

856:
861:            If arrayA(iterA - 1).dci_obr = arrayP(iterP - 1).dci_obr Then
862:                arraycruz(iterador).dci_obr = True
863:            Else
864:                arraycruz(iterador).dci_obr = False
865:            End If

922:            If arrayA(iterA - 1).cnpem = arrayP(iterP - 1).cnpem Then

923:                arraycruz(iterador).cnpem = True
924:            Else
925:                arraycruz(iterador).cnpem = False
926:            End If
                '927 a 931 não estava no 6.0 lite
927:            If arrayA(iterA - 1).nDCI = arrayP(iterP - 1).nDCI Then
928:                arraycruz(iterador).nDCI = True
929:            Else
930:                arraycruz(iterador).nDCI = False
931:            End If

93:             If arraycruz(iterador).cnpem = True Then
94:                 If arrayP(iterP - 1).cnpem = 0 Then
95:
96:                     arraycruz(iterador).xcnpem = 3 'sem cnpem
97:                 Else
98:                     arraycruz(iterador).xcnpem = 4 'cnpem é igual
99:                 End If
100:            Else 'é diferente
101:                If arrayP(iterP - 1).cnpem <> 0 And arrayA(iterA - 1).cnpem <> 0 Then 'se ambos <> 0
102:                    arraycruz(iterador).xcnpem = 0
103:                ElseIf arrayP(iterP - 1).cnpem = 0 Then 'se prescrito = 0
104:                    arraycruz(iterador).xcnpem = 1
105:                Else 'aviado =0
106:                    arraycruz(iterador).xcnpem = 2
107:                End If
108:            End If
                '109 a 131 não estavam  no 6.0 lite
109:            If arraycruz(iterador).GH = True Then
110:                If arrayP(iterP - 1).GH = 0 Then
111:                    arraycruz(iterador).xGH = 3 'sem GH
112:                Else
113:                    arraycruz(iterador).xGH = 4 'GH é igual
114:                End If
115:            Else 'é diferente
116:                If arrayP(iterP - 1).GH <> 0 And arrayA(iterA - 1).GH <> 0 Then 'se ambos <> 0
117:                    arraycruz(iterador).xGH = 0
118:                ElseIf arrayP(iterP - 1).GH = 0 Then 'se prescrito = 0
119:                    arraycruz(iterador).xGH = 1
120:                Else 'aviado =0
121:                    arraycruz(iterador).xGH = 2
122:                End If
123:            End If
124:
125:            If arraycruz(iterador).nDCI = True Then  'é true se ambos iguais
126:                If arrayP(iterP - 1).nDCI = False Then 'se pelo menos um é falso então é falso
127:                    arraycruz(iterador).xnDCI = False
128:                Else 'se ambos são verdadeiros é verdadeiro
129:                    arraycruz(iterador).xnDCI = True
130:                End If
131:            End If

133:            If arrayP(iterP - 1).code > 50000000 Then 'prescrito por CNPEM
134:                Select Case arraycruz(iterador).xcnpem
                        Case Is = 0 'mesmo cnpem
136:                        arraycruz(iterador).porDCImesmoCNPEM = True
137:                        arraycruz(iterador).porDCIdifCNPEM = False
138:                    Case Is = 1 'prescrito sem cnpem
139:                        arraycruz(iterador).porDCImesmoCNPEM = False
140:                        arraycruz(iterador).porDCIdifCNPEM = True
141:                    Case Is = 2 'aviado sem cnpem
142:                        arraycruz(iterador).porDCImesmoCNPEM = False
143:                        arraycruz(iterador).porDCIdifCNPEM = True
144:                    Case Is = 3 'ambos sem cnpem
145:                        arraycruz(iterador).porDCImesmoCNPEM = True
146:                        arraycruz(iterador).porDCIdifCNPEM = False
147:                    Case Is = 4 'cnpem diferente
148:                        arraycruz(iterador).porDCImesmoCNPEM = False
149:                        arraycruz(iterador).porDCIdifCNPEM = True
150:                        End Select
                    arraycruz(iterador).excepa = 5
                    arraycruz(iterador).excepb = 5
                    arraycruz(iterador).excepc = 5
                    arraycruz(iterador).porDCImesmoCNPEM = 5
151:            Else 'prescrito por marca ou lab (pode ser genérico)
152:                Select Case arraycruz(iterador).xcnpem
                        Case Is = 0 'mesmo cnpem
154:                        arraycruz(iterador).porMARCAmesmoCNPEM = True
155:                        arraycruz(iterador).porMARCAdifCNPEM = False
156:                    Case Is = 1 'prescrito sem cnpem
157:                        arraycruz(iterador).porMARCAmesmoCNPEM = False
158:                        arraycruz(iterador).porMARCAdifCNPEM = True
159:                    Case Is = 2 'aviado sem cnpem
160:                        arraycruz(iterador).porMARCAmesmoCNPEM = False
161:                        arraycruz(iterador).porMARCAdifCNPEM = True
162:                    Case Is = 3 'ambos sem cnpem
163:                        arraycruz(iterador).porMARCAmesmoCNPEM = True
164:                        arraycruz(iterador).porMARCAdifCNPEM = False
165:                    Case Is = 4 'cnpem diferente
166:                        arraycruz(iterador).porMARCAmesmoCNPEM = False
167:                        arraycruz(iterador).porMARCAdifCNPEM = True
168:                        End Select
                    If arrayP(iterP - 1).excepa = True Then 'se prescrito tem dci dos excep a
171:                    If arraycruz(iterador).code = False Then 'se prescrito e aviado têm codigo diferente
172:                        arraycruz(iterador).excepa = 0
173:                    Else
174:                        arraycruz(iterador).excepa = 1
175:                    End If
176:                Else 'se o dci não é dos excep a
177:                    arraycruz(iterador).excepa = 2
178:                End If
179:                If arrayP(iterP - 1).excepb = True Then 'se prescrito tem dci dos excep b
                        If arraycruz(iterador).code = False Then 'se prescrito e aviado têm codigo diferente
                            arraycruz(iterador).excepb = 0
                        Else
                            arraycruz(iterador).excepb = 1
                        End If
                    Else 'se o dci não é dos excep a
                        arraycruz(iterador).excepb = 2
                    End If

                    If arraycruz(iterador).code = False Then 'se prescrito e aviado têm codigo diferente
                        If arrayA(iterA - 1).pvpun > arrayP(iterP - 1).pvpun Then
                            arraycruz(iterador).excepc = 0 'excepc
                        ElseIf arrayA(iterA - 1).pvpun = arrayP(iterP - 1).pvpun Then
                            arraycruz(iterador).excepc = 1 'mp
                        ElseIf arrayA(iterA - 1).pvpun < arrayP(iterP - 1).pvpun Then
                            arraycruz(iterador).excepc = 2 'mb
                        End If
                    Else
                        arraycruz(iterador).excepc = 4 'certo
                    End If



                    If arrayA(iterA - 1).dci_obr = False Then 'Or arrayP(iterP - 1).nDCI = False Then 'não há genéricos 
                        'se nome diferente e código diferente =0 (tinha também o lab, tirei)
                        If arraycruz(iterador).nome = False And arraycruz(iterador).code = False Then 'And arraycruz(iterador).lab = False 
                            arraycruz(iterador).marcadifsemhavergens = 0
                        Else '=1
                            arraycruz(iterador).marcadifsemhavergens = 1
                        End If
                    Else 'há genéricos
                        arraycruz(iterador).marcadifsemhavergens = 2
                    End If
                   


                    If arrayP(iterP - 1).top5 > 0 Then 'prescrito tem top5
                        If arrayA(iterA - 1).pvp > arrayP(iterP - 1).top5 Then
                            arraycruz(iterador).top5 = 0
                        ElseIf arrayA(iterA - 1).pvp = arrayP(iterP - 1).top5 Then
                            arraycruz(iterador).top5 = 1
                        Else
                            arraycruz(iterador).top5 = 2
                        End If
                    Else
                        arraycruz(iterador).top5 = 3
                    End If


169:            End If
170:
180:
181:
                'End If
                iterP = iterP + 1
182:
            Loop


            'End If
183:        iterA = iterA + 1

184:    Loop
185:
186:
187:
188:
189:
190:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub comparador2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'vai buscar à baseDeDados as linhas correspondentes aos códigos introduzidos nas 8 caixas
    Sub irbuscar2013()
1:      On Error GoTo MOSTRARERRO
2:
3:      If op4.code > 0 Then
4:          PP = 4
5:      ElseIf op3.code > 0 Then
6:          PP = 3
7:      ElseIf op2.code > 0 Then
8:          PP = 2
9:      Else : PP = 1
10:     End If
11:
12:     If oa4.code > 0 Then
13:         AA = 4
14:     ElseIf oa3.code > 0 Then
15:         AA = 3
16:     ElseIf oa2.code > 0 Then
17:         AA = 2
18:     Else : AA = 1
19:     End If
20:
21:
22:
23:     If aviam4.Text <> "0" And aviam4.Text <> "" Then
24:         A = 4
25:         If aviam4.Text = " " Then
26:             A = 3
27:         End If
28:     ElseIf aviam3.Text <> "0" And aviam3.Text <> "" Then
29:         A = 3
30:         If aviam3.Text = " " Then
31:             A = 2
32:         End If
33:     ElseIf aviam2.Text <> "0" And aviam2.Text <> "" Then
34:         A = 2
35:         If aviam2.Text = " " Then
36:             A = 1
37:         End If
38:
39:     Else : A = 1
40:     End If
41:
42:     If presc4.Text <> "0" And presc4.Text <> "" Then
43:         P = 4
44:         If presc4.Text = " " Then
45:             P = 3
46:         End If
47:     ElseIf presc3.Text <> "0" And presc3.Text <> "" Then
48:         P = 3
49:         If presc3.Text = " " Then
50:             P = 2
51:         End If
52:     ElseIf presc2.Text <> "0" And presc2.Text <> "" Then
53:         P = 2
54:         If presc2.Text = " " Then
55:             P = 1
56:         End If
57:     Else : P = 1
58:     End If
59:
591:    Dim arrayoa(3) As Object
592:    arrayoa(0) = oa1
593:    arrayoa(1) = oa2
594:    arrayoa(2) = oa3
595:    arrayoa(3) = oa4
596:    Dim arrayop(3) As Object
597:    arrayop(0) = op1
598:    arrayop(1) = op2
599:    arrayop(2) = op3
600:    arrayop(3) = op4
601:    Dim arrayarow(3) As Object
602:    arrayarow(0) = a1row
603:    arrayarow(1) = a2row
604:    arrayarow(2) = a3row
605:    arrayarow(3) = a4row
606:    Dim arrayprow(3) As Object
607:    arrayprow(0) = p1row
608:    arrayprow(1) = p2row
609:    arrayprow(2) = p3row
610:    arrayprow(3) = p4row
611:    Dim arrayaarray(3) As Object
612:    arrayaarray(0) = a1array
613:    arrayaarray(1) = a2array
614:    arrayaarray(2) = a3array
615:    arrayaarray(3) = a4array
616:    Dim arrayparray(3) As Object
617:    arrayparray(0) = p1array
618:    arrayparray(1) = p2array
619:    arrayparray(2) = p3array
620:    arrayparray(3) = p4array
621:    Dim arrayaviado(3) As Object
622:    arrayaviado(0) = Aviado1
623:    arrayaviado(1) = Aviado2
624:    arrayaviado(2) = Aviado3
625:    arrayaviado(3) = Aviado4
626:    Dim arrayprescrito(3) As Object
627:    arrayprescrito(0) = Prescrito1
628:    arrayprescrito(1) = Prescrito2
629:    arrayprescrito(2) = Prescrito3
630:    arrayprescrito(3) = Prescrito4
631:
632:    For letraA As Integer = 0 To A - 1
633:        arrayarow(letraA) = DS.infarmed.FindBycode(arrayaviado(letraA).codigo)
634:        arrayaarray(letraA).Add(arrayarow(letraA))
635:        If Not IsNothing(arrayarow(letraA)) Then
636:            arrayoa(letraA).code = arrayarow(letraA)(0)
637:            arrayoa(letraA).nome = arrayarow(letraA)(2)
638:            arrayoa(letraA).d10910 = arrayarow(letraA)(17)
639:            arrayoa(letraA).pvpmenos2 = arrayarow(letraA)(24)
640:            arrayoa(letraA).cnpem = arrayarow(letraA)(25)
641:            arrayoa(letraA).dci = LCase(arrayarow(letraA)(1))
70:             arrayoa(letraA).forma = arrayarow(letraA)(3)
71:             arrayoa(letraA).dose = arrayarow(letraA)(4)
72:             arrayoa(letraA).qty = Replace(arrayarow(letraA)(5).ToString, ".", ",")
73:             arrayoa(letraA).comp = arrayarow(letraA)(6)
74:             arrayoa(letraA).gh = arrayarow(letraA)(7)
75:             arrayoa(letraA).gen = arrayarow(letraA)(10)
76:             arrayoa(letraA).lab = arrayarow(letraA)(11)
77:             arrayoa(letraA).pvp = arrayarow(letraA)(8)
78:             arrayoa(letraA).pr = arrayarow(letraA)(9)
79:             arrayoa(letraA).d4250 = arrayarow(letraA)(12)
80:             arrayoa(letraA).d1234 = arrayarow(letraA)(13)
81:             arrayoa(letraA).d10279 = arrayarow(letraA)(15)
82:             arrayoa(letraA).d10280 = arrayarow(letraA)(16)
83:             arrayoa(letraA).d21094 = arrayarow(letraA)(14)
84:             arrayoa(letraA).lei6 = arrayarow(letraA)(21)
85:             arrayoa(letraA).d14123 = arrayarow(letraA)(18)
86:             arrayoa(letraA).dci_obr = arrayarow(letraA)(20)
87:             arrayoa(letraA).trocamarca = arrayarow(letraA)(23)
88:             arrayoa(letraA).top5 = arrayarow(letraA)(19)
89:             arrayoa(letraA).pvpmenos1 = arrayarow(letraA)(22)
891:            If IsNumeric(arrayoa(letraA).qty) Then
892:                arrayoa(letraA).pvpun = arrayoa(letraA).pvp / arrayoa(letraA).qty
893:                arrayoa(letraA).pvpmenos1un = arrayoa(letraA).pvpmenos1 / arrayoa(letraA).qty
894:                arrayoa(letraA).pvpmenos2un = arrayoa(letraA).pvpmenos2 / arrayoa(letraA).qty
895:            Else
896:                arrayoa(letraA).pvpun = arrayoa(letraA).pvp
897:                arrayoa(letraA).pvpmenos1un = arrayoa(letraA).pvpmenos1
898:                arrayoa(letraA).pvpmenos2un = arrayoa(letraA).pvpmenos2
899:            End If
900:        End If
901:    Next
902:
903:
904:    For letraP As Integer = 0 To P - 1
905:        arrayprow(letraP) = DS.infarmed.FindBycode(arrayprescrito(letraP).codigo)
906:        arrayparray(letraP).Add(arrayprow(letraP))
907:        If Not IsNothing(arrayprow(letraP)) Then
908:            arrayop(letraP).code = arrayprow(letraP)(0)
909:            arrayop(letraP).nome = arrayprow(letraP)(2)
910:            arrayop(letraP).d10910 = arrayprow(letraP)(17)
911:            arrayop(letraP).pvpmenos2 = arrayprow(letraP)(24)
912:            arrayop(letraP).cnpem = arrayprow(letraP)(25)
913:            arrayop(letraP).dci = LCase(arrayprow(letraP)(1))
170:            arrayop(letraP).forma = arrayprow(letraP)(3)
171:            arrayop(letraP).dose = arrayprow(letraP)(4)
172:            arrayop(letraP).qty = Replace(arrayprow(letraP)(5).ToString, ".", ",")
173:            arrayop(letraP).comp = arrayprow(letraP)(6)
174:            arrayop(letraP).gh = arrayprow(letraP)(7)
175:            arrayop(letraP).gen = arrayprow(letraP)(10)
176:            arrayop(letraP).lab = arrayprow(letraP)(11)
177:            arrayop(letraP).pvp = arrayprow(letraP)(8)
178:            arrayop(letraP).pr = arrayprow(letraP)(9)
179:            arrayop(letraP).d4250 = arrayprow(letraP)(12)
180:            arrayop(letraP).d1234 = arrayprow(letraP)(13)
181:            arrayop(letraP).d10279 = arrayprow(letraP)(15)
182:            arrayop(letraP).d10280 = arrayprow(letraP)(16)
183:            arrayop(letraP).d21094 = arrayprow(letraP)(14)
184:            arrayop(letraP).lei6 = arrayprow(letraP)(21)
185:            arrayop(letraP).d14123 = arrayprow(letraP)(18)
186:            arrayop(letraP).dci_obr = arrayprow(letraP)(20)
187:            arrayop(letraP).trocamarca = arrayprow(letraP)(23)
188:            arrayop(letraP).top5 = arrayprow(letraP)(19)
189:            arrayop(letraP).pvpmenos1 = arrayprow(letraP)(22)
190:            If IsNumeric(arrayop(letraP).qty) Then
191:                arrayop(letraP).pvpun = arrayop(letraP).pvp / arrayop(letraP).qty
192:                arrayop(letraP).pvpmenos1un = arrayop(letraP).pvpmenos1 / arrayop(letraP).qty
193:                arrayop(letraP).pvpmenos2un = arrayop(letraP).pvpmenos2 / arrayop(letraP).qty
194:            Else
195:                arrayop(letraP).pvpun = arrayop(letraP).pvp
196:                arrayop(letraP).pvpmenos1un = arrayop(letraP).pvpmenos1
197:                arrayop(letraP).pvpmenos2un = arrayop(letraP).pvpmenos2
198:            End If
199:        End If
200:    Next
201:
233:
239:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub irbuscar2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub
    Dim rede As String

    Sub mostrarcodigo()
1:      On Error GoTo MOSTRARERRO

2:      Dim arrayresult(3) As Object
3:      arrayresult(0) = resultado1
4:      arrayresult(1) = resultado2
5:      arrayresult(2) = resultado3
6:      arrayresult(3) = resultado4

        Dim show(3) As embalagem.cruzamento  'assim além do ranking posso ir buscar o qualp e outras coisas
        show(0) = resultado1
        show(1) = resultado2
        show(2) = resultado3
        show(3) = resultado4
7:      Dim reresult(3) As Object
8:      reresult(0) = result1
9:      reresult(1) = result2
10:     reresult(2) = result3
11:     reresult(3) = result4
12:     Dim novoarray(44) As Object
13:     novoarray(11) = cruzamento11
        novoarray(12) = cruzamento12
        novoarray(13) = cruzamento13
        novoarray(14) = cruzamento14
        novoarray(21) = cruzamento21
        novoarray(22) = cruzamento22
        novoarray(23) = cruzamento23
        novoarray(24) = cruzamento24
        novoarray(31) = cruzamento31
        novoarray(32) = cruzamento32
        novoarray(33) = cruzamento33
        novoarray(34) = cruzamento34
        novoarray(41) = cruzamento41
        novoarray(42) = cruzamento42
        novoarray(43) = cruzamento43
        novoarray(44) = cruzamento44
        Dim cruzamentocerto(3) As Object
47:     cruzamentocerto(0) = resultado1
48:     cruzamentocerto(1) = resultado2
49:     cruzamentocerto(2) = resultado3
50:     cruzamentocerto(3) = resultado4
        Dim iterAAA As Single = 1
72:     Dim arrayoa(3) As Object
73:     arrayoa(0) = oa1
74:     arrayoa(1) = oa2
75:     arrayoa(2) = oa3
76:     arrayoa(3) = oa4
77:     Dim arrayop(3) As Object
78:     arrayop(0) = op1
79:     arrayop(1) = op2
80:     arrayop(2) = op3
81:     arrayop(3) = op4


        For toshow As Integer = 0 To A - 1
            Dim index As String = (show(toshow).qualP) & (toshow + 1)
18:         If Not IsNothing(novoarray(index)) Then
19:             If novoarray(index).cnpem = True Then
20:                 For secnpem As Integer = 1 To 5
                        If show(toshow).ranking.ToString.Substring(secnpem, 1) = 0 Or show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(secnpem, 1) = 2 Then 'só vai mudar se estiver a zero ou qty 1 ou 2
21:                         If secnpem = 4 Then
23:                             'marca não depende do CNPEM
25:                         ElseIf secnpem = 5 Then 'substitui o ranking da qty por certo no caso do cnpem ser o mesmo (para qty 0, 1 ou2)
26:                             show(toshow).xqty = show(toshow).ranking.ToString.Substring(secnpem, 1)
27:                             Mid(show(toshow).ranking, secnpem + 1, 1) = "3"
28:                         Else 'secnpem é 1, 2 ou 3 : dci, forma ou dose
29:                             'substitui o ranking da forma e dose por certo no caso do cnpem ser o mesmo
                                If arrayoa(toshow).lab <> 0 Then
31:                                 Mid(show(toshow).ranking, secnpem + 1, 1) = "1"
                                End If
32:                         End If
                        End If
33:                 Next
334:            End If
34:         End If
35:     Next


        Do While iterAAA <= A
            If aceitarduplicados = True Then
                'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
165:            For AAA As Integer = 1 To A
169:                If AAA <> iterAAA Then
                        If (arrayoa(AAA - 1).cnpem > 0 And (arrayoa(AAA - 1).cnpem = arrayoa(iterAAA - 1).cnpem)) Or (arrayoa(AAA - 1).code = arrayoa(iterAAA - 1).code) Then
170:                        If iterAAA > AAA Then 'para não fazer duas vezes e desfazer o que estava feito
                                cruzamentocerto(iterAAA - 1).ranking = cruzamentocerto(AAA - 1).ranking
                            End If
172:                    End If
173:                End If
174:            Next
            End If
            iterAAA = iterAAA + 1
        Loop

361:    For i As Integer = 0 To A - 1
            'MsgBox(arrayop(i).dci)
            'MsgBox(i + 1 & "º = " & arrayresult(i).ranking)
            For indice As Integer = 1 To 3
                If arrayresult(i).ranking.ToString.Substring(indice, 1) = 0 Then 'dci, ff, dose mal com cnpem diferente
                    'MsgBox("y, f ou g")
                    mal = True
                End If
            Next

378:        If (arrayresult(i).ranking.ToString.Substring(0, 1) = 0 Or arrayresult(i).ranking.ToString.Substring(0, 1) = 1 Or arrayresult(i).ranking.ToString.Substring(0, 1) = 2) AndAlso arrayresult(i).ranking.ToString.Substring(5, 1) = 2 Then 'L
379:            avisoL = True
                If arrayresult(i).ranking.ToString.Substring(4, 1) = 0 Then 'S
                    'MsgBox("S?" & i + 1)
                    If Not arrayresult(i).ranking.ToString.Substring(0, 1) = "0" Then 'para não mostrar marca dif quando prescrito por cnpem 1ºdig = 0
                        marcawarning = True
                        avisoL = False
                        'MsgBox("S" & i + 1)  

                    End If
                End If
380:            'MsgBox("L" & i + 1)
381:            'avisar para justificar embalagem mais pequena
            ElseIf arrayresult(i).ranking.ToString.Substring(0, 1) <> 0 AndAlso (arrayresult(i).ranking.ToString.Substring(5, 1) = 0 Or arrayresult(i).ranking.ToString.Substring(5, 1) = 1) Then 'h
374:            'MsgBox("h" & i + 1)
375:            avisoH = True
            ElseIf arrayresult(i).ranking.ToString.Substring(0, 1) = 0 AndAlso (arrayresult(i).ranking.ToString.Substring(5, 1) = 0 Or arrayresult(i).ranking.ToString.Substring(5, 1) = 1) Then 'h
3751:           'criei a partir do de cima porque quando prescrito por cnpem acusava marca dif em vez de h)
                'MsgBox("h" & i + 1)
3752:           avisoH = True
            ElseIf arrayresult(i).ranking.ToString.Substring(0, 1) = 0 Then
                'MsgBox("aqui")
                mal = True
            ElseIf arrayresult(i).ranking.ToString.Substring(4, 1) = 0 Then 'S
368:            'MsgBox("S?" & i + 1)
369:            If Not arrayresult(i).ranking.ToString.Substring(0, 1) = "0" Then 'para não mostrar marca dif quando prescrito por cnpem 1ºdig = 0
370:                marcawarning = True
                    'MsgBox("S" & i + 1)
371:            End If
382:        End If

364:        'if arrayresult(i).ranking.ToString.Substring(1, 3) = 111 Then
365:        'ok
366:
384:
385:    Next




        If mal = False And avisoL = False And avisoH = False And marcawarning = False Then
            codigoantes = Convert.ToInt32(getlatest())
            PictureBox1.Hide()
            resultadototal.Show()
            If showcode = True Then
                resultadototal.Font = New Font("calibri", 95.25)
                resultadototal.Font = New Font(resultadototal.Font, FontStyle.Bold)
                resultadototal.Text = (codigoantes + 1) & "." & A & reducao(A, codigoantes + 1)
            Else
                resultadototal.Text = "OK"
            End If
            cordefundo = Color.MediumSeaGreen
            PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
            If showcode = True Then
                'FileWriter = New StreamWriter("ofpacode", False)
                FileWriter = New StreamWriter(rede & "\ofpacode.", False)
14:             FileWriter.Write(codigoantes + 1)
15:             FileWriter.Close()
            End If

            For i As Integer = 0 To A - 1
                If arrayop(i).dci = "0" Then 'prescrito desconhecido
                    escrever(arrayop(i).code)
                End If
            Next

            resultadomostrado = True
            Me.Focus()
            Exit Sub
        End If

BAD:    If mal = True Then
            PictureBox1.Hide()
            resultadototal.Show()
            cordefundo = Color.Tomato
            resultadototal.Text = "ERRO"
            For i As Integer = 0 To A - 1
                If arrayresult(i).ranking = "000000" Then 'se há erro total de correspondência

                    'descobrir se erro está no presc ou no aviam
                    If arrayop(i).dci = "0" Then 'erro no prescrito" Then
                        'desconhecido = True
                        Select Case arrayop(i).porCNPEM
                            Case Is = True 'prescrito por CNPEM
                                justificacao.Text = arrayop(i).code & " desconhecido ou descontinuado"
                            Case Is = False 'prescrito por código AIM
                                justificacao.Text = arrayop(i).code & " desconhecido"
                        End Select
                        escrever(arrayop(i).code)

                    ElseIf arrayoa(i).dci = "0" Then 'erro no aviado
                        'desconhecido = True
                        justificacao.Text = arrayoa(i).code & " desconhecido"
                        escrever(arrayoa(i).code)
                    End If

                

                End If

            Next

22:         PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
            resultadomostrado = True
            Me.Focus()
            Exit Sub
        End If

WARNING: If mal = False And avisoL = True And avisoH = False Then
            codigoantes = Convert.ToInt32(getlatest())
            PictureBox1.Hide()
            resultadototal.Show()
            If showcode = True Then
                resultadototal.Font = New Font("arial", 50)
                resultadototal.Text = (codigoantes + 1) & "." & A + 5 & reducao(A, codigoantes + 1)
            Else
                resultadototal.Text = "OK"
            End If
            cordefundo = Color.Yellow
            PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            justificacao.BringToFront()
            justificacao.Text = "justificar caixa(s) menor(es)"
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
            If showcode = True Then
                'FileWriter = New StreamWriter("ofpacode", False)
                FileWriter = New StreamWriter(rede & "\ofpacode.", False)
                FileWriter.Write(codigoantes + 1)
                FileWriter.Close()
            End If
            resultadomostrado = True
            Me.Focus()
            Exit Sub
        End If

        If mal = False And avisoL = False And avisoH = True Then
            codigoantes = Convert.ToInt32(getlatest())
            PictureBox1.Hide()
            resultadototal.Show()
            If showcode = True Then
                resultadototal.Font = New Font("arial", 50)
                resultadototal.Text = (codigoantes + 1) & "." & A + 4 & reducao(A, codigoantes + 1)
            Else
                resultadototal.Text = "OK"
            End If
            'MsgBox(arrayresult(0).ranking)
            cordefundo = Color.Orange
            PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            justificacao.BringToFront()
            justificacao.Text = "justificar caixa(s) MAIOR(es)"
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
            If showcode = True Then
                'FileWriter = New StreamWriter("ofpacode", False)
                FileWriter = New StreamWriter(rede & "\ofpacode.", False)
                FileWriter.Write(codigoantes + 1)
                FileWriter.Close()
            End If
            resultadomostrado = True
            Me.Focus()
            Exit Sub
        End If

marcaWARNING: If mal = False And avisoL = False And avisoH = False And marcawarning = True Then
            PictureBox1.Hide()
            resultadototal.Show()
            cordefundo = Color.Tomato
            resultadototal.Text = "marca dif."

            PictureBox1.BackColor = cordefundo
            Me.BackColor = cordefundo
            resultadototal.BackColor = cordefundo
            justificacao.BackColor = cordefundo
            hora.BackColor = cordefundo
            result1.BackColor = cordefundo
            result2.BackColor = cordefundo
            result3.BackColor = cordefundo
            result4.BackColor = cordefundo
            resultadomostrado = True
            Me.Focus()
            Exit Sub
        End If



        resultadomostrado = True
        Me.Focus()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub mostrarcodigo: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub

    Function getlatest()
1:      On Error GoTo MOSTRARERRO
        'If File.Exists("ofpacode") Then
        If File.Exists(rede & "\ofpacode.") Then
2:          'FileReader = New StreamReader("ofpacode")
            FileReader = New StreamReader(rede & "\ofpacode.")
3:          getlatest = FileReader.ReadLine().ToString
            FileReader.Close()
        Else
            getlatest = 0
        End If
4:      Exit Function
MOSTRARERRO:
        MsgBox("function getlatest: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Dim s As String
    Dim n As Double
    Dim r As Double

    Public Function letrar(ByVal valor As Integer)
1:      On Error GoTo MOSTRARERRO
2:      letrar = Chr(64 + valor)
        Exit Function
MOSTRARERRO:
        MsgBox("function letrar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Public Function reducao(ByVal aviados As Integer, ByVal numero As Integer)
1:      On Error GoTo MOSTRARERRO
2:
3:      s = aviados & numero
4:      n = Val(s)
5:      While n > 26
6:
7:          s = 0
8:          While n > 0
9:
10:             r = n Mod 10
11:             s = s + r
12:             n = n \ 10
13:         End While
14:         n = s
15:     End While
16:
17:     reducao = letrar(s)
18:     Exit Function
MOSTRARERRO:
        MsgBox("function reducao: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Dim mesactual As Integer
    Dim mesguardado As Integer
    Sub month()
        On Error GoTo MOSTRARERRO
1:
2:      'If File.Exists("ofpames") Then
        If File.Exists(rede & "\ofpames.") Then
3:          'monthReader = New StreamReader("ofpames")
            monthReader = New StreamReader(rede & "\ofpames.")
4:          mesguardado = monthReader.ReadLine().ToString
5:          monthReader.Close()
6:      Else
7:          'monthWriter = New StreamWriter("ofpames", False)
            monthWriter = New StreamWriter(rede & "\ofpames.", False)
8:          monthWriter.Write(mesactual)
9:          monthWriter.Close()
            'FileWriter = New StreamWriter("ofpacode", False)
            FileWriter = New StreamWriter(rede & "\ofpacode.", False)
            FileWriter.Write(0)
            FileWriter.Close()
10:     End If
11:     If mesactual = mesguardado Then
            'do nothing
12:     Else
13:         'monthWriter = New StreamWriter("ofpames", False)
            monthWriter = New StreamWriter(rede & "\ofpames.", False)
14:         monthWriter.Write(mesactual)
15:         monthWriter.Close()
            'MessageBox.Show(mesactual)
16:         'FileWriter = New StreamWriter("ofpacode", False)
            FileWriter = New StreamWriter(rede & "\ofpacode.", False)
17:         FileWriter.Write(0)
18:         FileWriter.Close()
19:     End If
20:     Exit Sub
MOSTRARERRO:
        MsgBox("sub month: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Dim inicialguardado As String
    Dim PC As New dadosPC
    Function autenticacao()
1:      On Error GoTo MOSTRARERRO
2:
5:      'PC.sistema não me deixou no system32 nem no windows
6:      If File.Exists(PC.fileinicial) Then
7:          inicialReader = New StreamReader(PC.fileinicial)
8:          inicialguardado = inicialReader.ReadLine().ToString
9:          inicialReader.Close()
10:
11:         If inicialguardado = PC.nome Then
12:             autenticacao = True
13:         ElseIf inicialguardado = "orgafarma·pt" Then
14:             inicialWriter = New StreamWriter(PC.fileinicial, False)
17:             inicialWriter.Write(PC.nome)
18:             inicialWriter.Close()
19:             autenticacao = True
20:         Else
21:             MessageBox.Show("ficheiro corrompido", "PrescAvi v6.0 lite")
22:             autenticacao = False
23:         End If
24:     Else
25:         MessageBox.Show("falta ficheiro", "PrescAvi v6.0 lite")
26:         autenticacao = False
27:     End If
28:
29:     Exit Function
MOSTRARERRO:
        MsgBox("sub autenticacao: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Private Sub escrever(unknown As String)
        'FileWriter = New StreamWriter(rede & "\unknown.txt", True)
14:     'FileWriter.WriteLine(unknown)
15:     'FileWriter.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Select Case CheckBox1.Checked
            Case Is = True
                showcode = True
            Case Is = False
                showcode = False
        End Select
        inicializar()
        presc1.Focus()
    End Sub
End Class
