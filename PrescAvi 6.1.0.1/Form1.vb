Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
'parece que mudei os arrays para novos números mas não mudei os rows, só o 17 para 8 em 18/5/13. comparar com versão pré bd aumentada
'row(#); array(#)
'(0)code
'(1)dci
'(2)nome
'(3)forma
'(4)dose
'(5)qty
'(6)comp
'(7)GH
'(8)gen?
'(9)lab
'(10)4250
'(11)1234
'(12)21094
'(13)10279
'(14)10280
'(15)10910
'(16)14123
'(17)top5
'(18)dci_obr
'(19)lei6
'(20)trocamarca [importa o novo despacho de 2014 da ictiose]
'(21)pvpmenos2 (era pvp-3)
'(22)CNPEM
'(23)pvp
'(24)pr
'(25)pvpmenos1 (era pvpold)
'(26)pvpmenos3 (era pvp-4)
'(27)pvpmenos4 (era pvp-5)
'(28)pvpmenos5 (era pvp-6)

'MessageBox.Show(IWin32Window, String, String)
'Displays a message box in front of the specified object and with the specified text and caption.


'nDCI era assim:
'If  a1row(9) = 0 'sem obrigatoriedade de prescrição por dci
'    marcadif(7, 0)
'End If



'exemplo iterações
'Dim sb As New System.Text.StringBuilder
'
'       For i As Integer = 0 To results.Length - 1
'          sb.Append(results(i) & vbCrLf)
'     Next



'exemplo iterações
'For i As Integer = 0 To strOriginal.Length - 1
'           MessageBox.Show(strOriginal(i).ToString())
'    Next i






'exemplo iterações
'For index As Integer = 0 To SubTables.Count - 1
'Dim RefDesgStr() As String = Nothing
'Dim ii As Integer = 0
'
'     For Each row As DataRow In SubTables(index).Rows
'
'              RefDesgStr(ii) = row.Item("RefDesg")   ' Got the error here
'' some code checking RefDesgStr(ii) with all the prior strings in RefDesgStr()
'         ii +=1
'     Next row
'  index+=1
'next

Public Class Form1
    Inherits Form
    Dim naoindicar As Boolean
    Dim valorcombopvp1 As Boolean
    Dim valorcombopvp2 As Boolean
    Dim valorcombopvp3 As Boolean
    Dim valorcombopvp4 As Boolean

    Dim mesactual As Single = Date.Now.Month
    Dim concatenado As String
    Dim antesp(3) As Double
    Dim depoisp(3) As Double
    Dim antesa(3) As Double
    Dim depoisa(3) As Double
    Dim repeticao As New Integer
    Dim aceitarduplicados As Boolean = False
    Dim verificador As Color
    Dim erro As Boolean
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
    Dim ordenado1(3)
    Dim ordenado2(3)
    Dim ordenado3(3)
    Dim ordenado4(3)
    Dim inicializado As Boolean
    'Private _passedText As String

    'Public Property [PassedText]() As String
    '  Get
    '     Return _passedText
    'End Get
    'Set(ByVal Value As String)
    '    _passedText = Value
    'End Set
    '    End Property
    Dim comparado As Boolean
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
    'Dim varcruza5 As Short
    'Dim varcruza6 As Short
    Dim varcruzp1 As Short
    Dim varcruzp2 As Short
    Dim varcruzp3 As Short
    Dim varcruzp4 As Short
    'Dim varcruzp5 As Short
    'Dim varcruzp6 As Short
    Dim firsttime As Boolean = True
    Dim varlabelcruz1 As String
    Dim varlabelcruz2 As String
    Dim varlabelcruz3 As String
    Dim varlabelcruz4 As String
    'Dim varlabelcruz5 As String
    'Dim varlabelcruz6 As String
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
    'Dim portcomp5 As Double
    'Dim portcomp6 As Double
    Dim intermedio As Double
    Dim intermedio2 As Double
    Dim pvp1v As String
    Dim pvp2v As String
    Dim pvp3v As String
    Dim pvp4v As String
    'Dim pvp5v As String
    'Dim pvp6v As String
    Dim pvp1val As Double
    Dim pvp2val As Double
    Dim pvp3val As Double
    Dim pvp4val As Double
    'Dim pvp5val As Double
    'Dim pvp6val As Double
    Dim comp1v As String
    Dim comp2v As String
    Dim comp3v As String
    Dim comp4v As String
    'Dim comp5v As String
    'Dim comp6v As String
    Dim comp1val As Double
    Dim comp2val As Double
    Dim comp3val As Double
    Dim comp4val As Double
    Dim pvp11v As String
    Dim pvp22v As String
    Dim pvp33v As String
    Dim pvp44v As String
    'Dim pvp5v As String
    'Dim pvp6v As String
    Dim pvp11val As Double
    Dim pvp22val As Double
    Dim pvp33val As Double
    Dim pvp44val As Double
    'Dim pvp5val As Double
    'Dim pvp6val As Double
    Dim comp11v As String
    Dim comp22v As String
    Dim comp33v As String
    Dim comp44v As String
    'Dim comp5v As String
    'Dim comp6v As String
    Dim comp11val As Double
    Dim comp22val As Double
    Dim comp33val As Double
    Dim comp44val As Double
    'Dim comp5val As Double
    'Dim comp6val As Double
    Dim pr As Double
    Dim pr1 As Double
    Dim pr2 As Double
    Dim pr3 As Double
    Dim pr4 As Double
    Dim pr11 As Double
    Dim pr22 As Double
    Dim pr33 As Double
    Dim pr44 As Double
    'Dim pr5 As Double
    'Dim pr6 As Double
    Dim organismo As String
    Dim gen As Boolean
    Dim comp As Double
    Dim compdois As Double
    Dim portimedio As String
    'Dim p5row As basededadosDataSet.infarmedRow
    'Dim p6row As basededadosDataSet.infarmedRow
    'Dim a5row As basededadosDataSet.infarmedRow
    'Dim a6row As basededadosDataSet.infarmedRow

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

    Dim form2 As New bd()

    Dim p1, p2, p3, p4, a1, a2, a3, a4 As Integer

    'trabalhar a data da comparticipação - no exemplo 40023 bate certo com 29/07/2009 - é só substituir por aXarray(6)
    'Dim descomp As Date = DateAdd(DateInterval.Day, 40023, #12/30/1899#)
    'em vez de abrir a msgbox comparar logo com data actual e se fora de data dar como não comparticiapado (¿desde...?)
    Dim descomp1mostrado As Boolean
    Dim descomp2mostrado As Boolean
    Dim descomp3mostrado As Boolean
    Dim descomp4mostrado As Boolean

    Dim Prescrito1 As New meds
    Dim Prescrito2 As New meds
    Dim Prescrito3 As New meds
    Dim Prescrito4 As New meds
    'Dim Prescrito5 As New meds
    'Dim Prescrito6 As New meds

    Dim Aviado1 As New meds
    Dim Aviado2 As New meds
    Dim Aviado3 As New meds
    Dim Aviado4 As New meds
    'Dim Aviado5 As New meds
    'Dim Aviado6 As New meds
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



    'Dim umTA As New ficheiro1DataSetTableAdapters.umTableAdapter
    'Dim FS As New ficheiro1DataSet
    'Dim umTAs As Integer = umTA.Fill(FS.um)



    Dim a1array As New ArrayList
    Dim a2array As New ArrayList
    Dim a3array As New ArrayList
    Dim a4array As New ArrayList
    'Dim a5array As New ArrayList
    'Dim a6array As New ArrayList
    Dim p1array As New ArrayList
    Dim p2array As New ArrayList
    Dim p3array As New ArrayList
    Dim p4array As New ArrayList
    'Dim p5array As New ArrayList
    'Dim p6array As New ArrayList

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
    'Dim av5 As New avaliacao
    'Dim av6 As New avaliacao

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
    Dim resultadomostrado As Boolean

    'Prepara para ler nova receita - limpa tudo (valores a zero, labels em branco e sem fundo) e foco na primeira caixa
    Sub inicializar()

        On Error GoTo MOSTRARERRO
        naoindicar = True
        top5Box.SelectionAlignment = HorizontalAlignment.Center
        excepBox.SelectionAlignment = HorizontalAlignment.Center
        PictureBox2.Hide()
        PictureBox1.Show()
1000:   'Me.BackColor = SystemColors.GradientInactiveCaption
1001:   '  For Each C As Control In Controls
1002:   '   If TypeOf C Is Label Then
1003:   '   DirectCast(C, Label).BackColor = SystemColors.GradientInactiveCaption
1004:   '   End If
1005:   '  If TypeOf C Is RichTextBox Then
1006:   '   DirectCast(C, RichTextBox).BackColor = SystemColors.GradientInactiveCaption
1007:   '  End If
        ' If TypeOf C Is GroupBox Then
        'DirectCast(C, GroupBox).BackColor = SystemColors.GradientInactiveCaption
        'End If
1008:   ' If C.HasChildren Then
1009:   '       UpdateLabelFG(C.Controls, SystemColors.GradientInactiveCaption)
1010:   '      End If
1011:   'Next
11101:  ' hora.BackColor = SystemColors.GradientActiveCaption
11102:
11103:  If valorcombopvp1 = True Then
11104:      ComboBox1.DataSource = Nothing
11105:      ComboBox1.DataBindings.Clear()
11106:  End If
11107:
11108:  If valorcombopvp2 = True Then
11109:      ComboBox2.DataSource = Nothing
11110:      ComboBox2.DataBindings.Clear()
11111:  End If
11112:
11113:  If valorcombopvp3 = True Then
11114:      ComboBox3.DataSource = Nothing
11115:      ComboBox3.DataBindings.Clear()
11116:  End If
11117:
11118:  If valorcombopvp4 = True Then
11119:      ComboBox4.DataSource = Nothing
11120:      ComboBox4.DataBindings.Clear()
11121:  End If
11122:
11123:  valorcombopvp1 = False
11124:  valorcombopvp2 = False
11125:  valorcombopvp3 = False
11126:  valorcombopvp4 = False
11127:  concatenado = "00"
11128:  repeticao = 0
11129:  Array.Clear(antesp, 0, 4)
11130:  Array.Clear(antesa, 0, 4)
11131:  Array.Clear(depoisp, 0, 4)
11132:  Array.Clear(depoisa, 0, 4)
11133:  C_PVP_1.Text = ""
11134:  C_PVP_2.Text = ""
11135:  C_PVP_3.Text = ""
11136:  C_PVP_4.Text = ""
11137:  C_PVP_1.BackColor = SystemColors.GradientInactiveCaption
11138:  C_PVP_2.BackColor = SystemColors.GradientInactiveCaption
11139:  C_PVP_3.BackColor = SystemColors.GradientInactiveCaption
11140:  C_PVP_4.BackColor = SystemColors.GradientInactiveCaption
11141:  top5Box.Text = ""
11142:  excepBox.Text = ""
11143:  resultadomostrado = False
11144:  read_only()
11145:  verificador = Color.Green
11146:  erro = False
11147:  comparado = False





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
        resultado1.erro = False
        resultado2.erro = False
        resultado3.erro = False
        resultado4.erro = False
        resultado1.qualP = 0
        resultado2.qualP = 0
        resultado3.qualP = 0
        resultado4.qualP = 0
        resultado1.ranking = "0000000000"
        resultado2.ranking = "0000000000"
        resultado3.ranking = "0000000000"
        resultado4.ranking = "0000000000"
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
        cruzamento11.ranking = 0
        cruzamento12.ranking = 0
        cruzamento13.ranking = 0
        cruzamento14.ranking = 0
        cruzamento21.ranking = 0
        cruzamento22.ranking = 0
        cruzamento23.ranking = 0
        cruzamento24.ranking = 0
        cruzamento31.ranking = 0
        cruzamento32.ranking = 0
        cruzamento33.ranking = 0
        cruzamento34.ranking = 0
        cruzamento41.ranking = 0
        cruzamento42.ranking = 0
        cruzamento43.ranking = 0
        cruzamento44.ranking = 0

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



        op1.pvpmenos3 = vbEmpty
        op1.pvpmenos4 = vbEmpty
        op1.pvpmenos5 = vbEmpty
        op1.pvpexcepC = vbEmpty
        op2.pvpmenos3 = vbEmpty
        op2.pvpmenos4 = vbEmpty
        op2.pvpmenos5 = vbEmpty
        op2.pvpexcepC = vbEmpty
        op3.pvpmenos3 = vbEmpty
        op3.pvpmenos4 = vbEmpty
        op3.pvpmenos5 = vbEmpty
        op3.pvpexcepC = vbEmpty
        op4.pvpmenos3 = vbEmpty
        op4.pvpmenos4 = vbEmpty
        op4.pvpmenos5 = vbEmpty
        op4.pvpexcepC = vbEmpty

        vshe = 0
        vshe1 = 0
        vshe2 = 0
        vshe3 = 0
        vshe4 = 0
        sehaport.Text = ""
        limpo = True
        agrupado = False

        but1474_01.Visible = False
        but1234_01.Visible = False
        but14123_01.Visible = False
        but4250_01.Visible = False
        but21094_01.Visible = False
        but10910_01.Visible = False
        but10279_01.Visible = False
        but62010_01.Visible = False
        but1474_02.Visible = False
        but1234_02.Visible = False
        but14123_02.Visible = False
        but4250_02.Visible = False
        but21094_02.Visible = False
        but10910_02.Visible = False
        but10279_02.Visible = False
        but62010_02.Visible = False
        but1474_03.Visible = False
        but1234_03.Visible = False
        but14123_03.Visible = False
        but4250_03.Visible = False
        but21094_03.Visible = False
        but10910_03.Visible = False
        but10279_03.Visible = False
        but62010_03.Visible = False
        but1474_04.Visible = False
        but1234_04.Visible = False
        but14123_04.Visible = False
        but4250_04.Visible = False
        but21094_04.Visible = False
        but10910_04.Visible = False
        but10279_04.Visible = False
        but62010_04.Visible = False

a:      varlabelcruz1 = ""
b:      labelcruz1.Text = ""
c:      varlabelcruz2 = ""
d:      labelcruz2.Text = ""
e:      varlabelcruz3 = ""
f:      labelcruz3.Text = ""
g:      varlabelcruz4 = ""
h:      labelcruz4.Text = ""
i:      'varlabelcruz5 = ""
j:      'labelcruz5.Text = ""
k:      'varlabelcruz6 = ""
l:      'labelcruz6.Text = ""
1:      av1.nivel = 100
2:      av2.nivel = 100
3:      av3.nivel = 100
4:      av4.nivel = 100
5:      a1_4250 = False
6:      a2_4250 = False
7:      a3_4250 = False
8:      a4_4250 = False
9:      a1_1234 = False
10:     a2_1234 = False
11:     a3_1234 = False
12:     a4_1234 = False
13:     a1_14123 = False
14:     a2_14123 = False
15:     a3_14123 = False
16:     a4_14123 = False
17:     a1_21094 = False
18:     a2_21094 = False
19:     a3_21094 = False
20:     a4_21094 = False
21:     a1_1474ad = False
22:     a2_1474ad = False
23:     a3_1474ad = False
24:     a4_1474ad = False
25:     a1_1474nl = False
26:     a2_1474nl = False
27:     a3_1474nl = False
28:     a4_1474nl = False
29:     a1_10279 = False
30:     a2_10279 = False
31:     a3_10279 = False
32:     a4_10279 = False
33:     a1_10910 = False
34:     a2_10910 = False
35:     a3_10910 = False
36:     a4_10910 = False
37:     descomp1mostrado = False
38:     descomp2mostrado = False
39:     descomp3mostrado = False
40:     descomp4mostrado = False
41:     A = 0
42:     P = 0
43:
47:     Me.presc1.BackColor = Color.White
48:     Me.presc2.BackColor = Color.White
49:     Me.presc3.BackColor = Color.White
50:     Me.presc4.BackColor = Color.White
51:     Me.aviam1.BackColor = Color.White
52:     Me.aviam2.BackColor = Color.White
53:     Me.aviam3.BackColor = Color.White
54:     Me.aviam4.BackColor = Color.White
55:     Me.aviam1.Text = ""
56:     Me.aviam2.Text = ""
57:     Me.aviam3.Text = ""
58:     Me.aviam4.Text = ""
59:     Me.presc1.Text = ""
60:     Me.presc2.Text = ""
61:     Me.presc3.Text = ""
62:     Me.presc4.Text = ""
63:     Me.result1.Text = ""
64:     Me.result2.Text = ""
65:     Me.result3.Text = ""
66:     Me.result4.Text = ""
67:     Me.result1.BackColor = SystemColors.GradientInactiveCaption
68:     Me.result2.BackColor = SystemColors.GradientInactiveCaption
69:     Me.result3.BackColor = SystemColors.GradientInactiveCaption
70:     Me.result4.BackColor = SystemColors.GradientInactiveCaption
71:     Me.presc1.Focus()
72:     Me.Prescrito1.codigo = Nothing
73:     Me.Prescrito2.codigo = Nothing
74:     Me.Prescrito3.codigo = Nothing
75:     Me.Prescrito4.codigo = Nothing
76:     Me.Aviado1.codigo = Nothing
77:     Me.Aviado2.codigo = Nothing
78:     Me.Aviado3.codigo = Nothing
79:     Me.Aviado4.codigo = Nothing
80:     Me.a1p1.nivel = 100
81:     Me.a2p1.nivel = 100
82:     Me.a3p1.nivel = 100
83:     Me.a4p1.nivel = 100
84:     Me.a1p2.nivel = 100
85:     Me.a2p2.nivel = 100
86:     Me.a3p2.nivel = 100
87:     Me.a4p2.nivel = 100
88:     Me.a1p3.nivel = 100
89:     Me.a2p3.nivel = 100
90:     Me.a3p3.nivel = 100
91:     Me.a4p3.nivel = 100
92:     Me.a1p4.nivel = 100
93:     Me.a2p4.nivel = 100
94:     Me.a3p4.nivel = 100
95:     Me.a4p4.nivel = 100
96:     Me.a1p1.resultado = 100
97:     Me.a2p1.resultado = 100
98:     Me.a3p1.resultado = 100
99:     Me.a4p1.resultado = 100
100:    Me.a1p2.resultado = 100
101:    Me.a2p2.resultado = 100
102:    Me.a3p2.resultado = 100
103:    Me.a4p2.resultado = 100
104:    Me.a1p3.resultado = 100
105:    Me.a2p3.resultado = 100
106:    Me.a3p3.resultado = 100
107:    Me.a4p3.resultado = 100
108:    Me.a1p4.resultado = 100
109:    Me.a2p4.resultado = 100
110:    Me.a3p4.resultado = 100
111:    Me.a4p4.resultado = 100
112:    Me.a1p1.mostrado = False
113:    Me.a2p1.mostrado = False
114:    Me.a3p1.mostrado = False
115:    Me.a4p1.mostrado = False
116:    Me.a1p2.mostrado = False
117:    Me.a2p2.mostrado = False
118:    Me.a3p2.mostrado = False
119:    Me.a4p2.mostrado = False
120:    Me.a1p3.mostrado = False
121:    Me.a2p3.mostrado = False
122:    Me.a3p3.mostrado = False
123:    Me.a4p3.mostrado = False
124:    Me.a1p4.mostrado = False
125:    Me.a2p4.mostrado = False
126:    Me.a3p4.mostrado = False
127:    Me.a4p4.mostrado = False
128:
132:    'presc1gen.Checked = False
133:    'presc2gen.Checked = False
134:    'presc3gen.Checked = False
135:    'presc4gen.Checked = False
136:    grupoP1 = 0
137:    grupoP2 = 0
138:    grupoP3 = 0
139:    grupoP4 = 0
140:    grupoA1 = 0
141:    grupoA2 = 0
142:    grupoA3 = 0
143:    grupoA4 = 0
144:    grupoP1dci = ""
145:    grupoP2dci = ""
146:    grupoP3dci = ""
147:    grupoP4dci = ""
148:    grupoA1dci = ""
149:    grupoA2dci = ""
150:    grupoA3dci = ""
151:    grupoA4dci = ""
152:    varcruza1 = 0
153:    varcruza2 = 0
154:    varcruza3 = 0
155:    varcruza4 = 0
156:    'varcruza5 = 0
157:    'varcruza6 = 0
158:    varcruzp1 = 0
159:    varcruzp2 = 0
160:    varcruzp3 = 0
161:    varcruzp4 = 0
162:
174:    Me.port1.Text = ""
175:    Me.port2.Text = ""
176:    Me.port3.Text = ""
177:    Me.port4.Text = ""
178:    'Me.port5.Text = ""
179:    'Me.port6.Text = ""
180:    Me.pvp1.Text = ""
181:    Me.pvp2.Text = ""
182:    Me.pvp3.Text = ""
183:    Me.pvp4.Text = ""
        Me.pvp11.Text = ""
        Me.pvp22.Text = ""
        Me.pvp33.Text = ""
        Me.pvp44.Text = ""
        'Me.pvp5.Text = ""
185:    'Me.pvp6.Text = ""
186:    Me.comp1.Text = ""
187:    Me.comp2.Text = ""
188:    Me.comp3.Text = ""
189:    Me.comp4.Text = ""
        Me.comp11.Text = ""
        Me.comp22.Text = ""
        Me.comp33.Text = ""
        Me.comp44.Text = ""
190:    'Me.comp5.Text = ""
191:    'Me.comp6.Text = ""
192:    Me.totalPVP.Text = ""
193:    Me.totalComp.Text = ""
        Me.totalPVP2.Text = ""
        Me.totalComp2.Text = ""
194:    tirarports()
195:    verde = True
196:    amarelo = False
197:    vermelho = False
198:    verifgenlabel.Text = ""
199:    verifgenlabel.BackColor = Color.Transparent
19901:  conjunto = 0
19902:
19903:  Select Case mesactual
            Case Is = 1, 4, 7, 10
19905:          'usam-se 4 (o actual e os 3 anteriores)
19906:          oa1.arraypvp = {"", "", "", ""}
19907:          oa2.arraypvp = {"", "", "", ""}
19908:          oa3.arraypvp = {"", "", "", ""}
19909:          oa4.arraypvp = {"", "", "", ""}
19910:      Case Is = 2, 5, 8, 11
19911:          'usam-se 5 (o actual e os 4 anteriores)
19912:          oa1.arraypvp = {"", "", "", "", ""}
19913:          oa2.arraypvp = {"", "", "", "", ""}
19914:          oa3.arraypvp = {"", "", "", "", ""}
19915:          oa4.arraypvp = {"", "", "", "", ""}
19916:      Case Is = 3, 6, 9, 12
19917:          'usam-se 6 (o actual e os 5 anteriores)
19918:          oa1.arraypvp = {"", "", "", "", "", ""}
19919:          oa2.arraypvp = {"", "", "", "", "", ""}
19920:          oa3.arraypvp = {"", "", "", "", "", ""}
19921:          oa4.arraypvp = {"", "", "", "", "", ""}
19922:          End Select
19923:  Select Case mesactual
            Case Is = 1, 4, 7, 10
19925:          'usam-se 4 (o actual e os 3 anteriores)
19926:          Array.Clear(oa1.arraypvp, 0, 4)
19927:          Array.Clear(oa2.arraypvp, 0, 4)
19928:          Array.Clear(oa3.arraypvp, 0, 4)
19929:          Array.Clear(oa4.arraypvp, 0, 4)
19930:      Case Is = 2, 5, 8, 11
19931:          'usam-se 5 (o actual e os 4 anteriores)
19932:          Array.Clear(oa1.arraypvp, 0, 5)
19933:          Array.Clear(oa2.arraypvp, 0, 5)
19934:          Array.Clear(oa3.arraypvp, 0, 5)
19935:          Array.Clear(oa4.arraypvp, 0, 5)
19936:      Case Is = 3, 6, 9, 12
19937:          'usam-se 6 (o actual e os 5 anteriores)
19938:          Array.Clear(oa1.arraypvp, 0, 6)
19939:          Array.Clear(oa2.arraypvp, 0, 6)
19940:          Array.Clear(oa3.arraypvp, 0, 6)
19941:          Array.Clear(oa4.arraypvp, 0, 6)
19942:          End Select
19943:
19944:
19945:
19946:
19947:
19948:  ComboBox1.Refresh()
19949:  ComboBox2.Refresh()
19950:  ComboBox3.Refresh()
19951:  ComboBox4.Refresh()
19952:
19953:  firsttime = False
19954:  naoindicar = False
19955:  Exit Sub
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

    Dim exepath As String
    Dim bdpath As String
    'o que acontece ao abrir/inicializar o form principal - inicia timer, limpa tudo, poe tudo a zero e foca na 1ª textbox
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'BasededadosDataSet1.infarmed' table. You can move, or remove it, as needed.
        On Error GoTo MOSTRARERRO
        bdpath = "basededados.mdb"
        Me.Text = Application.ProductName & " " & Application.ProductVersion & " (" & FileDateTime(bdpath).ToString.Substring(6, 4) & FileDateTime(bdpath).ToString.Substring(3, 2) & FileDateTime(bdpath).ToString.Substring(0, 2) & ")"
        'farmacia = Convert.ToInt32(_passedText)
        MarcaMarcaToolStripMenuItem.Checked = True
        TrocaDeLaboratórioToolStripMenuItem.Checked = False
        Me.InfarmedTableAdapter1.Fill(Me.BasededadosDataSet1.infarmed)


1:      'TODO: This line of code loads data into the 'EscolhaDataSet22.lab' table. You can move, or remove it, as needed.
2:      ' Me.LabTableAdapter3.Fill(Me.EscolhaDataSet22.lab)
3:      'TODO: This line of code loads data into the 'EscolhaDataSet21.lab' table. You can move, or remove it, as needed.
4:      ' Me.LabTableAdapter2.Fill(Me.EscolhaDataSet21.lab)
5:      'TODO: This line of code loads data into the 'EscolhaDataSet20.lab' table. You can move, or remove it, as needed.
6:      ' Me.LabTableAdapter.Fill(Me.EscolhaDataSet20.lab)
7:      'TODO: This line of code loads data into the 'EscolhaDataSet19.qty' table. You can move, or remove it, as needed.
8:      ' Me.QtyTableAdapter3.Fill(Me.EscolhaDataSet19.qty)
9:      'TODO: This line of code loads data into the 'EscolhaDataSet18.qty' table. You can move, or remove it, as needed.
10:     ' Me.QtyTableAdapter2.Fill(Me.EscolhaDataSet18.qty)
11:     'TODO: This line of code loads data into the 'EscolhaDataSet17.qty' table. You can move, or remove it, as needed.
12:     'Me.QtyTableAdapter.Fill(Me.EscolhaDataSet17.qty)
13:     'TODO: This line of code loads data into the 'EscolhaDataSet16.forma' table. You can move, or remove it, as needed.
14:     'Me.FormaTableAdapter3.Fill(Me.EscolhaDataSet16.forma)
15:     ''TODO: This line of code loads data into the 'EscolhaDataSet15.forma' table. You can move, or remove it, as needed.
16:     'Me.FormaTableAdapter2.Fill(Me.EscolhaDataSet15.forma)
17:     'TODO: This line of code loads data into the 'EscolhaDataSet14.forma' table. You can move, or remove it, as needed.
18:     'Me.FormaTableAdapter.Fill(Me.EscolhaDataSet14.forma)
19:     'TODO: This line of code loads data into the 'EscolhaDataSet13.dc1' table. You can move, or remove it, as needed.
20:     'Me.Dc1TableAdapter4.Fill(Me.EscolhaDataSet13.dc1)
21:     'TODO: This line of code loads data into the 'EscolhaDataSet12.dc1' table. You can move, or remove it, as needed.
22:     'Me.Dc1TableAdapter3.Fill(Me.EscolhaDataSet12.dc1)
23:     'TODO: This line of code loads data into the 'EscolhaDataSet11.dc1' table. You can move, or remove it, as needed.
24:     'Me.Dc1TableAdapter2.Fill(Me.EscolhaDataSet11.dc1)
25:     'TODO: This line of code loads data into the 'EscolhaDataSet10.dc1' table. You can move, or remove it, as needed.
26:     ' Me.Dc1TableAdapter1.Fill(Me.EscolhaDataSet10.dc1)
27:     'TODO: This line of code loads data into the 'EscolhaDataSet9.lab' table. You can move, or remove it, as needed.
28:     'Me.LabTableAdapter1.Fill(Me.EscolhaDataSet9.lab)
29:     'TODO: This line of code loads data into the 'EscolhaDataSet8.qty' table. You can move, or remove it, as needed.
30:     'Me.QtyTableAdapter1.Fill(Me.EscolhaDataSet8.qty)
31:     'TODO: This line of code loads data into the 'EscolhaDataSet7.dose' table. You can move, or remove it, as needed.
32:     'Me.DoseTableAdapter1.Fill(Me.EscolhaDataSet7.dose)
33:     'TODO: This line of code loads data into the 'EscolhaDataSet6.forma' table. You can move, or remove it, as needed.
34:     ' Me.FormaTableAdapter1.Fill(Me.EscolhaDataSet6.forma)
35:     'TODO: This line of code loads data into the 'EscolhaDataSet5.dc1' table. You can move, or remove it, as needed.
36:     'Me.Dc1TableAdapter.Fill(Me.EscolhaDataSet5.dc1)
37:     '   Me.InfarmedTableAdapter.Fill(Me.BasededadosDataSet.infarmed)
38:     Timer1.Start()

39:     inicializar()

        Me.WindowState = FormWindowState.Maximized
40:     Me.KeyPreview = True
        'but01.Checked = True
        'organismo = "01"
        'organismotxt.Text = organismo
        'farmaciar(farmacia)
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Form1_Load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub data_Keyspace(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'no lite o space está no keysEnter
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
            inicializar()
            limpar4()
            limpar41()
            CodeEC1.Text = ""
            codEC2.Text = ""
178:        Me.Focus()
179:        inicializar()
            Me.presc1.Text = ""
        End If
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

        If e.KeyCode = Keys.Enter Then
            Select Case foco
                Case "codeEC1"
41:                 If CodeEC1.Text = "" Then
42:                     'Beep()
43:                     'My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
44:                 Else

45:                     Me.codEC2.Focus()
46:                     'codEC2.SelectionStart = 0
47:                     'codEC2.SelectionLength = Len(codEC2.Text)
48:
49:                     'codigo4.codigo = codEC2.Text
50:                     'incorporar()
51:
52:                     'mostrar()
53:                 End If


                Case "codEC2"
                    If codEC2.Text = "" Then
                        'Beep()
                        'My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
                    Else
                        'codEC2.Text = ""
                        'CodeEC1.Text = ""
                        'limpar4()
                        'limpar41()
                        Me.CodeEC1.Focus()
                        ' CodeEC1.SelectionStart = 0
                        'CodeEC1.SelectionLength = Len(CodeEC1.Text)
                        'codigo41.codigo = CodeEC1.Text
                        'incorporar()
                        'mostrar()
                    End If




54:             Case "presc1"
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
                    'Case "presc5"
z77:                '   If presc5.Text = "" Then
z78:                '                Me.aviam1.Focus()
zz80:               '               Me.presc5.Text = "0"
z81:                '              Me.presc6.Text = "0"
zz81:               '             Else
z82:                '            Me.aviam1.Focus()
                    '           Me.presc6.Text = "0"
z83:                '          End If
z84:                '     Case "presc6"
z85:                '        If presc6.Text = "" Then
z86:                '                Me.aviam1.Focus()
z87:                '               Me.presc6.Text = "0"
z88:                '              Else
z89:                '             Me.aviam1.Focus()
z90:                '            End If
91:             Case "aviam2"
92:                 If aviam2.Text = "" Then
93:                     Me.aviam2.Text = "0"
94:                     Me.aviam3.Text = "0"
95:                     Me.aviam4.Text = "0"
a95:                    '             Me.aviam5.Text = "0"
b95:                    '            Me.aviam6.Text = "0"
96:                     Comparar()
a96:                    somas()
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
a103:                   '           Me.aviam5.Text = "0"
b103:                   '          Me.aviam6.Text = "0"
104:                    Comparar()
a104:                   somas()
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
z94:                    '         Me.aviam5.Text = "0"
z95:                    '        Me.aviam6.Text = "0"
                        Comparar()

zzz96:                  somas()

                        Me.presc1.Focus()
z96:                Else

                        '       Me.aviam5.Text = "0"
                        '      Me.aviam6.Text = "0"
b110:                   If IsNumeric(aviam4.Text) Then
                            If aviam4.Text >= 1111111 And aviam4.Text <= 9999999 Then
c110:                           Comparar()
d110:                           somas()
e110:                           Me.presc1.Focus()
f110:                       Else
g110:                           'Beep()
h110:                           Me.aviam4.Text = ""
i110:                           Me.aviam4.Focus()
z99:                        End If
                        Else
                            aviam4.Text = ""

                            Me.aviam4.Focus()
                        End If
                    End If


z100:               'Case "aviam5"
z101:               '   If aviam5.Text = "" Then
z102:               '               Me.aviam5.Text = "0"
z103:               '              Me.aviam6.Text = "0"
z104:               '             Comparar()
                    '            somas()
                    '           Me.presc1.Focus()
z105:               '          Else
                    '         Me.aviam6.Text = "0"
                    '        If aviam5.Text >= 1111111 And aviam5.Text <= 9999999 Then
a108:               '               Comparar()
b108:               '              somas()
c108:               '             Me.presc1.Focus()
d108:               '            Else
e108:               '           Beep()
f108:               '          Me.aviam5.Text = ""
g108:               '         Me.aviam5.Focus()
h108:               '        End If
                    '       End If
108:                '  Case "aviam6"
109:                '     If aviam6.Text = "" Then
110:                '                Me.aviam6.Text = "0"
                    '               Comparar()
                    '              somas()
                    '             Me.presc1.Focus()
a111:               '            Else
b111:               '           If aviam6.Text >= 1111111 And aviam6.Text <= 9999999 Then
c111:               '               Comparar()
d111:               '              somas()
e111:               '             Me.presc1.Focus()
f111:               '            Else
g111:               '           Beep()
                    '          Me.aviam6.Text = ""
h111:               '         Me.aviam6.Focus()
i111:               '        End If
j111:               '       End If
112:                'Comparar()
                    'somas()
113:                End Select
        End If
    End Sub

    Private Sub data_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Add Then
            e.SuppressKeyPress = True
            limpar4()
            limpar41()
            codEC2.Text = ""
            Me.CodeEC1.Focus()
10:         CodeEC1.SelectionStart = 0
11:         CodeEC1.SelectionLength = Len(CodeEC1.Text)
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




    'mostrar o form da bd ao clickar no botão - retirado quando fiz menu principal
    'Private Sub formBD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles formBD.Click
    '   On Error GoTo MOSTRARERRO
    '  form2.Show()
    ' Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub formBD_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub

    'mostrar o form da bd ao fazer enter no botão - retirado quando fiz menu principal
    ' Private Sub formBD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles formBD.Enter
    '    On Error GoTo MOSTRARERRO
    '   form2.Show()
    '  Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub formBD_Enter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub



    'o que acontece ao fazer enter no botão iniciar - lança inicializar
    'Private Sub iniciar_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   On Error GoTo MOSTRARERRO
    '  limpar4()
    '2:      codEC2.Text = ""
    '1:      inicializar()
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Sub iniciar_Enter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '  End Sub



    'o que acontece ao clickar no botão iniciar - lança inicializar
    '    Private Sub iniciar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        On Error GoTo MOSTRARERRO
    '1:      'antigamente tinha isto: Me.aviam1.Clear()
    '2:      inicializar()
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Sub iniciar_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub


    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        On Error GoTo MOSTRARERRO
        If e.KeyCode = Keys.F1 Then
3:          inicializar()
4:      End If

        If e.KeyCode = Keys.F12 Then
127:        Comparar()
128:    End If
129:
130:    If e.KeyCode = Keys.F2 Then
131:        form2.Show()
132:    End If
133:
134:    If e.KeyCode = Keys.Control AndAlso e.KeyCode = Keys.I Then
135:        inicializar()
136:    End If
137:
138:    If e.KeyCode = Keys.Control AndAlso e.KeyCode = Keys.C Then
139:        Comparar()
140:    End If
141:
142:    If e.KeyCode = Keys.Control AndAlso e.KeyCode = Keys.B Then
143:        form2.Show()
144:    End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Form1_KeyDown: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
11:             Beep()
12:
        End Select

35:
        'retirei o "c" e o "C" como enter e foram substituidos por $m ou $J com full ascii - senão fazia shift, apareciam simbolos e andava para trás
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


    'Private Sub presc5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc5.TextChanged
    '   On Error GoTo MOSTRARERRO
    '  Dim Caractp5 As String
    ' Caractp5 = presc5.Text
    'If semcod5.Checked = False Then
    '   If Len(presc5.Text) = 7 Then
    '      If Caractp5 Like "#######" Then
    '         Prescrito5.codigo = presc5.Text
    'op5.code = presc5.Text
    '    op5.CNPEM = False
    '        'Me.presc6.Focus()
    '   End If
    '  End If
    'Else
    '    presc5.Text = "0"
    'End If
    'Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub presc5_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub

    'Private Sub presc6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    On Error GoTo MOSTRARERRO
    '    Dim Caractp6 As String
    '   Caractp6 = presc6.Text
    '   If semcod6.Checked = False Then
    ' Len(presc6.Text) = 7 Then
    '           If Caractp6 Like "#######" Then
    '               Prescrito6.codigo = presc6.Text
    ' op6.code = presc6.Text
    '           op6.CNPEM = False
    '               'Me.aviam1.Focus()
    '           End If
    '       End If
    '   Else
    '       presc6.Text = "0"
    '   End If
    '   Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub presc6_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub



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
                'indicar(1)
                'indicar2(1)
                'Me.av2.Focus()
6:              'Me.aviam2.Focus()
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
                'indicar(2)
                'indicar2(2)
                'Me.av3.Focus()
6:              'Me.aviam3.Focus()
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
                'indicar(3)
                'indicar2(3)
                'Me.av4.Focus()
6:              'Me.aviam4.Focus()
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
                'indicar(4)
                'indicar2(4)
                'Me.av5.Focus()
6:              'MsgBox("para compensar o enter da caneta")
7:              'Comparar()
8:          End If
9:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviam4_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    ' Private Sub aviam5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam5.TextChanged
    '     On Error GoTo MOSTRARERRO
    '     Dim Caracta5 As String
    '     Caracta5 = aviam5.Text
    '    If Len(aviam5.Text) = 7 Then
    '        If Caracta5 Like "#######" Then
    '            Aviado5.codigo = aviam5.Text
    'oa5.code = aviam5.Text
    '            av5.mostrado = "True"
    '            A = 5
    '            a5row = DS.infarmed.FindBycode(Aviado5.codigo)
    '            codigorow = DS.infarmed.FindBycode(Aviado5.codigo)
    '            a5array.Add(a5row)
    '            indicar(5)
    'indicar2(5)
    '            'Me.av6.Focus()
    '            'MsgBox("para compensar o enter da caneta")
    '            'Comparar()
    '        End If
    '    End If
    '    Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub aviam5_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub


    'Private Sub aviam6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    On Error GoTo MOSTRARERRO
    '    Dim Caracta6 As String
    '    Caracta6 = aviam6.Text
    '    If Len(aviam6.Text) = 7 Then
    '        If Caracta6 Like "#######" Then
    '            Aviado6.codigo = aviam6.Text
    ' oa6.code = aviam6.Text
    '            av6.mostrado = "True"
    '            A = 6
    '            a6row = DS.infarmed.FindBycode(Aviado6.codigo)
    '            codigorow = DS.infarmed.FindBycode(Aviado6.codigo)
    '            a6array.Add(a6row)
    '            indicar(6)
    'indicar2(6)
    '            'Me.p1.Focus()
    '            'MsgBox("para compensar o enter da caneta")
    '            'Comparar()
    '        End If
    '   End If
    '   Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Sub aviam6_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '  End Sub








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

    'vai buscar à baseDeDados as linhas correspondentes aos códigos introduzidos nas 8 caixas


    Public Sub atribuirp() 'não está a ser chamado - não serve para nada
        On Error GoTo MOSTRARERRO
1:      If P >= 4 Then
            Prescrito4.principio = p4row(1)
3:          Prescrito4.apresentacao = p4row(3)
4:          Prescrito4.dosagem = p4row(4)
5:          Prescrito4.quantidade = p4row(5)
6:          Prescrito4.comparticipacao = p4row(6)
7:          Prescrito4.grupo = p4row(7)
8:          Prescrito4.generico = p4row(8)
9:          Prescrito4.laboratorio = p4row(9)
10:     ElseIf P >= 3 Then
            Prescrito4 = vazio
12:         Prescrito3.principio = p3row(1)
13:         Prescrito3.apresentacao = p3row(3)
14:         Prescrito3.dosagem = p3row(4)
15:         Prescrito3.quantidade = p3row(5)
16:         Prescrito3.comparticipacao = p3row(6)
17:         Prescrito3.grupo = p3row(7)
18:         Prescrito3.generico = p3row(8)
19:         Prescrito3.laboratorio = p3row(9)
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
30:         Prescrito2.generico = p2row(8)
31:         Prescrito2.laboratorio = p2row(9)
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
43:         Prescrito1.generico = p1row(8)
44:         Prescrito1.laboratorio = p1row(9)
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
8:          Aviado4.generico = a4row(8)
9:          Aviado4.laboratorio = a4row(9)
10:     ElseIf A >= 3 Then
11:         Aviado4 = vazio
            Aviado3.principio = a3row(1)
13:         Aviado3.apresentacao = a3row(3)
14:         Aviado3.dosagem = a3row(4)
15:         Aviado3.quantidade = a3row(5)
16:         Aviado3.comparticipacao = a3row(6)
17:         Aviado3.grupo = a3row(7)
18:         Aviado3.generico = a3row(8)
19:         Aviado3.laboratorio = a3row(9)
20:     ElseIf A >= 2 Then
21:         Aviado4 = vazio
22:         Aviado3 = vazio
            Aviado2.principio = a2row(1)
24:         Aviado2.apresentacao = a2row(3)
25:         Aviado2.dosagem = a2row(4)
26:         Aviado2.quantidade = a2row(5)
27:         Aviado2.comparticipacao = a2row(6)
28:         Aviado2.grupo = a2row(7)
29:         Aviado2.generico = a2row(8)
30:         Aviado2.laboratorio = a2row(9)
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
41:         Aviado1.generico = a1row(8)
42:         Aviado1.laboratorio = a1row(9)
43:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub atribuira: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Sub cruz(ByVal varcruzp As Short, ByVal avava As Short)
1:      On Error GoTo MOSTRARERRO
2:      Select Case avava
            Case Is = 1
3:              varcruza1 = varcruzp
4:          Case Is = 2
5:              varcruza2 = varcruzp
6:          Case Is = 3
7:              varcruza3 = varcruzp
8:          Case Is = 4
9:              varcruza4 = varcruzp
10:
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("sub cruz: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub mostrarcruz()
1:      On Error GoTo MOSTRARERRO
2:      varlabelcruz1 = "[" & varcruza1 & "]" & " -> [1]"
3:      varlabelcruz2 = "[" & varcruza2 & "]" & " -> [2]"
4:      varlabelcruz3 = "[" & varcruza3 & "]" & " -> [3]"
5:      varlabelcruz4 = "[" & varcruza4 & "]" & " -> [4]"
6:      'varlabelcruz5 = "[" & varcruza5 & "]" & " -> [5]"
7:      'varlabelcruz6 = "[" & varcruza6 & "]" & " -> [6]"
14:     If aviam2.Text = "0" Or aviam2.Text = "" Or result2.Text = "não aviado" Then
15:         varlabelcruz2 = ""
16:     End If
17:     If aviam3.Text = "0" Or aviam3.Text = "" Or result3.Text = "não aviado" Then
18:         varlabelcruz3 = ""
19:     End If
20:     If aviam4.Text = "0" Or aviam4.Text = "" Or result4.Text = "não aviado" Then
21:         varlabelcruz4 = ""
22:     End If
25:     'If aviam5.Text = "0" Or aviam5.Text = "" Or result5.Text = "não aviado" Then
26:     ':     varlabelcruz5 = ""
27:     '     End If
28:     '    If aviam6.Text = "0" Or aviam6.Text = "" Or result6.Text = "não aviado" Then
29:     '     varlabelcruz6 = ""
30:     '    End If
31:     labelcruz1.Text = varlabelcruz1
32:     labelcruz2.Text = varlabelcruz2
33:     labelcruz3.Text = varlabelcruz3
34:     labelcruz4.Text = varlabelcruz4
35:     '   labelcruz5.Text = varlabelcruz5
36:     '  labelcruz6.Text = varlabelcruz6



        Exit Sub
MOSTRARERRO:
        MsgBox("Sub mostrarcruz: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Function genlab(ByVal pgen As Boolean, ByVal agen As Boolean, ByVal plab As String, ByVal alab As String) As Short
        On Error GoTo MOSTRARERRO
        If filtrolab = True Then
1:          Select Case agen
                Case False

3:                  If pgen = False Then
4:                      If alab = plab Then

13:                         Return 0
14:                     Else
15:                         Return 5
16:                     End If

25:                 Else

26:                     Return 2
27:                 End If
28:             Case True
29:                 If pgen = False Then
30:                     Return 3

31:                 Else

32:                     If alab = plab Then
33:                         Return 1

34:                     Else
35:                         Return 4

36:                     End If
37:                 End If
38:                 End Select
        Else
            Return 0
        End If
        Exit Function
MOSTRARERRO:
        MsgBox("Function GenLab: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Sub OK(ByVal qual As Short, h As Short)
        On Error GoTo MOSTRARERRO
1:
        Select Case h
            Case 0 'qty =
                Select Case qual
                    Case 1
3:                      result1.Text = "OK"
4:                      result1.BackColor = Color.Green
5:                  Case 2
6:                      result2.Text = "OK"
7:                      result2.BackColor = Color.Green
8:                  Case 3
9:                      result3.Text = "OK"
10:                     result3.BackColor = Color.Green
11:                 Case 4
12:                     result4.Text = "OK"
13:                     result4.BackColor = Color.Green
                        ' Case 5
14:                     '    result5.Text = "OK"
15:                     '   result5.BackColor = Color.Green
16:                     'Case 6
17:                     '   result6.Text = "OK"
18:                     '  result6.BackColor = Color.Green

19:                     End Select
            Case 1 'qty >
                vermelho = True
                Select Case qual
                    Case 1
23:                     Hquant(1)
25:                 Case 2
26:                     Hquant(2)
28:                 Case 3
29:                     Hquant(3)
                    Case 4
                        Hquant(4)
                        'Case 5
                        '   Hquant(5)
                        'Case 6
                        '   Hquant(6)
                End Select
            Case 2 'qty <=
                Select Case qual
                    Case 1
                        result1.Text = "OK"
                        result1.BackColor = Color.Green
                    Case 2
                        result2.Text = "OK"
                        result2.BackColor = Color.Green
                    Case 3
                        result3.Text = "OK"
                        result3.BackColor = Color.Green
                    Case 4
                        result4.Text = "OK"
                        result4.BackColor = Color.Green
                        ' Case 5
                        '    result5.Text = "OK"
                        '   result5.BackColor = Color.Green
                        'Case 6
                        '   result6.Text = "OK"
                        '  result6.BackColor = Color.Green
                End Select
            Case 3 'qty ñ numerico
                Select Case qual
                    Case 1
                        verifQuant(1)
                    Case 2
                        verifQuant(2)
                    Case 3
                        verifQuant(3)
                    Case 4
                        verifQuant(4)
                        'Case 5
                        '   verifQuant(5)
                        'Case 6
                        '   verifQuant(6)
                End Select
            Case 9 'não interessa
                Select Case qual
                    Case 1
                        result1.Text = "OK"
                        result1.BackColor = Color.Green
                    Case 2
                        result2.Text = "OK"
                        result2.BackColor = Color.Green
                    Case 3
                        result3.Text = "OK"
                        result3.BackColor = Color.Green
                    Case 4
                        result4.Text = "OK"
                        result4.BackColor = Color.Green
                        'Case 5
                        '   result5.Text = "OK"
                        '   result5.BackColor = Color.Green
                        'Case 6
                        '   result6.Text = "OK"
                        '  result6.BackColor = Color.Green

                End Select
        End Select
        Exit Sub

MOSTRARERRO:
        MsgBox("Sub OK: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub Hquant(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "h) qty > 150%   ou  CNPEM"
4:              result1.BackColor = Color.Red
5:          Case 2
6:              result2.Text = "h) qty > 150%   ou  CNPEM"
7:              result2.BackColor = Color.Red
8:          Case 3
9:              result3.Text = "h) qty > 150%   ou  CNPEM"
10:             result3.BackColor = Color.Red
11:         Case 4
12:             result4.Text = "h) qty > 150%   ou  CNPEM"
13:             result4.BackColor = Color.Red
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Hquant: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub DoseDif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.BackColor = Color.Red
4:              result1.Text = "G) dosagem diferente da prescrita"
5:          Case 2
6:              result2.BackColor = Color.Red
7:              result2.Text = "G) dosagem diferente da prescrita"
8:          Case 3
9:              result3.BackColor = Color.Red
10:             result3.Text = "G) dosagem diferente da prescrita"
11:         Case 4
12:             result4.BackColor = Color.Red
13:             result4.Text = "G) dosagem diferente da prescrita"
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub DoseDif: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub dciDif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
                If result1.Text <> "" Or result1.Text <> " " Or result1.Text <> "0" Then
3:                  result1.BackColor = Color.Red
4:                  result1.Text = "y) DCI não prescrito"
                End If
5:          Case 2
                If result2.Text <> "" Or result2.Text <> " " Or result2.Text <> "0" Then
6:                  result2.BackColor = Color.Red
7:                  result2.Text = "y) DCI não prescrito"
                End If
8:          Case 3
                If result3.Text <> "" Or result3.Text <> " " Or result3.Text <> "0" Then
9:                  result3.BackColor = Color.Red
10:                 result3.Text = "y) DCI não prescrito"
                End If
11:         Case 4
                If result4.Text <> "" Or result4.Text <> " " Or result4.Text <> "0" Then
12:                 result4.BackColor = Color.Red
13:                 result4.Text = "y) DCI não prescrito"
                End If
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub dciDif: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub apresDif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.BackColor = Color.Red
4:              result1.Text = "f) apresentação diferente da prescrita"
5:          Case 2
6:              result2.BackColor = Color.Red
7:              result2.Text = "f) apresentação diferente da prescrita"
8:          Case 3
9:              result3.BackColor = Color.Red
10:             result3.Text = "f) apresentação diferente da prescrita"
11:         Case 4
12:             result4.BackColor = Color.Red
13:             result4.Text = "f) apresentação diferente da prescrita"
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub apresDif: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Sub nPresc(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
        vermelho = True
1:      If A > P Then
2:          Select Case qual
                Case 1
4:                  result1.Text = "->y) embalagem não prescrita"
5:                  result1.BackColor = Color.Red
6:              Case 2
7:                  result2.Text = "->y) embalagem não prescrita"
8:                  result2.BackColor = Color.Red
9:              Case 3
10:                 result3.Text = "->y) embalagem não prescrita"
11:                 result3.BackColor = Color.Red
12:             Case 4
13:                 result4.Text = "->y) embalagem não prescrita"
14:                 result4.BackColor = Color.Red
15:                 End Select
16:     Else
17:         Select Case qual
                Case 1
19:                 result1.Text = "y) DCI não prescrito"
20:                 result1.BackColor = Color.Red
21:             Case 2
22:                 result2.Text = "y) DCI não prescrito"
23:                 result2.BackColor = Color.Red
24:             Case 3
25:                 result3.Text = "y) DCI não prescrito"
26:                 result3.BackColor = Color.Red
27:             Case 4
28:                 result4.Text = "y) DCI não prescrito"
29:                 result4.BackColor = Color.Red
30:                 End Select
31:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub nPresc: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub marca(ByVal qual As Short, ByVal h As Short, ByVal para As String)
        On Error GoTo MOSTRARERRO
        If filtrolab = True Then
            vermelho = True
            Select Case h
                Case Is = 0 'qty =
1:                  Select Case qual
                        Case 1
3:                          result1.Text = "Aviado de marca (" & para & ") !!!"
4:                          result1.BackColor = Color.Orange
5:                      Case 2
6:                          result2.Text = "Aviado de marca (" & para & ") !!!"
7:                          result2.BackColor = Color.Orange
8:                      Case 3
9:                          result3.Text = "Aviado de marca (" & para & ") !!!"
10:                         result3.BackColor = Color.Orange
11:                     Case 4
12:                         result4.Text = "Aviado de marca (" & para & ") !!!"
13:                         result4.BackColor = Color.Orange
                            'Case 5
14:                         '    result5.Text = "Aviado de marca (" & para & ") !!!"
15:                         '    result5.BackColor = Color.Orange
                            'Case 6
16:                         '    result6.Text = "Aviado de marca (" & para & ") !!!"
17:                         '    result6.BackColor = Color.Orange
18:                         End Select
19:             Case Is = 1 'qty >
20:                 Select Case qual
                        Case 1
23:                         result1.Text = "h) + Aviado de marca (" & para & ")"
24:                         result1.BackColor = Color.Red
25:                     Case 2
26:                         result2.Text = "h) + Aviado de marca (" & para & ")"
27:                         result2.BackColor = Color.Red
28:                     Case 3
29:                         result3.Text = "h) + Aviado de marca (" & para & ")"
30:                         result3.BackColor = Color.Red
31:                     Case 4
32:                         result4.Text = "h) + Aviado de marca (" & para & ")"
33:                         result4.BackColor = Color.Red
                            'Case 5
34:                         '    result5.Text = "h) + Aviado de marca (" & para & ")"
35:                         '    result5.BackColor = Color.Red
                            'Case 6
36:                         '    result6.Text = "h) + Aviado de marca (" & para & ")"
37:                         '    result6.BackColor = Color.Red
38:                         End Select


                Case Is = 2 'qty <=
                    Select Case qual
                        Case Is = 1
                            result1.Text = "Aviado de marca (" & para & ") !!!"
                            result1.BackColor = Color.Red
                        Case Is = 2
                            result2.Text = "Aviado de marca (" & para & ") !!!"
                            result2.BackColor = Color.Red
                        Case Is = 3
                            result3.Text = "Aviado de marca (" & para & ") !!!"
                            result3.BackColor = Color.Red
                        Case Is = 4
                            result4.Text = "Aviado de marca (" & para & ") !!!"
                            result4.BackColor = Color.Red
                            ' Case Is = 5
                            '     result5.Text = "Aviado de marca (" & para & ") !!!"
                            '     result5.BackColor = Color.Red
                            ' Case Is = 6
                            '     result6.Text = "Aviado de marca (" & para & ") !!!"
                            '     result6.BackColor = Color.Red
                    End Select

                Case Is = 3 'qty <> numerico
                    Select Case qual
                        Case Is = 1
                            result1.Text = "Aviado de marca (" & para & ") !!!"
                            result1.BackColor = Color.Red
                            '    labelf1.Text = "h?"
                            '    labelf1.BackColor = Color.BlueViolet
                        Case Is = 2
                            result2.Text = "Aviado de marca (" & para & ") !!!"
                            result2.BackColor = Color.Red
                            '   labelf2.Text = "h?"
                            '  labelf2.BackColor = Color.BlueViolet
                        Case Is = 3
                            result3.Text = "Aviado de marca (" & para & ") !!!"
                            result3.BackColor = Color.Red
                            ' labelf3.Text = "h?"
                            'labelf3.BackColor = Color.BlueViolet
                        Case Is = 4
                            result4.Text = "Aviado de marca (" & para & ") !!!"
                            result4.BackColor = Color.Red
                            'labelf4.Text = "h?"
                            'labelf4.BackColor = Color.BlueViolet
                            '     Case Is = 5
                            '        result5.Text = "Aviado de marca (" & para & ") !!!"
                            '       result5.BackColor = Color.Red
                            '      labelf5.Text = "h?"
                            '     labelf5.BackColor = Color.BlueViolet
                            'Case Is = 6
                            '   result6.Text = "Aviado de marca (" & para & ") !!!"
                            '  result6.BackColor = Color.Red
                            ' labelf6.Text = "h?"
                            'labelf6.BackColor = Color.BlueViolet
                    End Select
            End Select
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub marca: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub marca2gen(ByVal qual As Short, ByVal h As Short)
        On Error GoTo MOSTRARERRO
        If filtrolab = True Then
            Select Case h
                Case Is = 0 'qty=
                    amarelo = True
1:                  Select Case qual
                        Case 1
3:                          'result1.Text = "ver se há desautorização [Marca -> " & LCase(a1array(1)) & "]"
                            result1.Text = "ver se há desautorização [Marca -> genérico]"
4:                          result1.BackColor = Color.Yellow
5:                      Case 2
6:                          'result2.Text = "ver se há desautorização [Marca -> " & LCase(a2array(1)) & "]"
                            result2.Text = "ver se há desautorização [Marca -> genérico]"
7:                          result2.BackColor = Color.Yellow
8:                      Case 3
9:                          'result3.Text = "ver se há desautorização [Marca -> " & LCase(a3array(1)) & "]"
                            result3.Text = "ver se há desautorização [Marca -> genérico]"
10:                         result3.BackColor = Color.Yellow
11:                     Case 4
12:                         'result4.Text = "ver se há desautorização [Marca -> " & LCase(a4array(1)) & "]"
                            result4.Text = "ver se há desautorização [Marca -> genérico]"
13:                         result4.BackColor = Color.Yellow
                            ' Case 5
19:                         '     'result5.Text = "ver se há desautorização [Marca -> " & LCase(a5array(1)) & "]"
                            '     result5.Text = "ver se há desautorização [Marca -> genérico]"
20:                         '     result5.BackColor = Color.Yellow
21:                         ' Case 6
22:                         '     'result6.Text = "ver se há desautorização [Marca -> " & LCase(a6array(1)) & "]"
                            '     result6.Text = "ver se há desautorização [Marca -> genérico]"
23:                         '     result6.BackColor = Color.Yellow
24:                         End Select
                Case Is = 1 'qty >
                    vermelho = True
                    Select Case qual
                        Case 1
33:                         'result1.Text = "h) + ver se há desautorização [Marca -> " & LCase(a1array(1)) & "]"
                            result1.Text = "h) + ver se há desautorização [Marca -> genérico]"
34:                         result1.BackColor = Color.Red
35:                     Case 2
36:                         'result2.Text = "h) + ver se há desautorização [Marca -> " & LCase(a2array(1)) & "]"
                            result2.Text = "h) + ver se há desautorização [Marca -> genérico]"
37:                         result2.BackColor = Color.Red
38:                     Case 3
39:                         'result3.Text = "h) + ver se há desautorização [Marca -> " & LCase(a3array(1)) & "]"
                            result3.Text = "h) + ver se há desautorização [Marca -> genérico]"
40:                         result3.BackColor = Color.Red
41:                     Case 4
42:                         'result4.Text = "h) + ver se há desautorização [Marca -> " & LCase(a4array(1)) & "]"
                            result4.Text = "h) + ver se há desautorização [Marca -> genérico]"
43:                         result4.BackColor = Color.Red
                            'Case 5
49:                         '    'result5.Text = "h) + ver se há desautorização [Marca -> " & LCase(a5array(1)) & "]"
                            '    result5.Text = "h) + ver se há desautorização [Marca -> genérico]"
50:                         '    result5.BackColor = Color.Red
51:                         'Case 6
52:                         '    'result6.Text = "h) + ver se há desautorização [Marca -> " & LCase(a6array(1)) & "]"
                            '    result6.Text = "h) + ver se há desautorização [Marca -> genérico]"
53:                         '    result6.BackColor = Color.Red
54:                         End Select


                Case Is = 2 'qty <=
                    amarelo = True
                    Select Case qual

                        Case 1
63:                         'result1.Text = "ver se há desautorização [Marca -> " & LCase(a1array(1)) & "]"
                            result1.Text = "ver se há desautorização [Marca -> genérico]"
64:                         result1.BackColor = Color.Yellow
65:                     Case 2
66:                         'result2.Text = "ver se há desautorização [Marca -> " & LCase(a2array(1)) & "]"
                            result2.Text = "ver se há desautorização [Marca -> genérico]"
67:                         result2.BackColor = Color.Yellow
68:                     Case 3
69:                         'result3.Text = "ver se há desautorização [Marca -> " & LCase(a3array(1)) & "]"
                            result3.Text = "ver se há desautorização [Marca -> genérico]"
70:                         result3.BackColor = Color.Yellow
71:                     Case 4
72:                         'result4.Text = "ver se há desautorização [Marca -> " & LCase(a4array(1)) & "]"
                            result4.Text = "ver se há desautorização [Marca -> genérico]"
73:                         result4.BackColor = Color.Yellow
                            'Case 5
79:                         '    'result5.Text = "ver se há desautorização [Marca -> " & LCase(a5array(1)) & "]"
                            '    result5.Text = "ver se há desautorização [Marca -> genérico]"
80:                         '    result5.BackColor = Color.Yellow
81:                         'Case 6
82:                         'result6.Text = "ver se há desautorização [Marca -> " & LCase(a6array(1)) & "]"
                            '    result6.Text = "ver se há desautorização [Marca -> genérico]"
83:                         '    result6.BackColor = Color.Yellow
                    End Select

                Case Is = 3 'qty não numerico
                    amarelo = True
                    Select Case qual
                        Case Is = 1
                            amarelo = True
                            'result1.Text = "ver se há desautorização [Marca -> " & LCase(a1array(1)) & "]"
                            result1.Text = "ver se há desautorização [Marca -> genérico]"
                            result1.BackColor = Color.Yellow
                            ' labellab1.Text = "h?"
                            ' labellab1.BackColor = Color.BlueViolet
                        Case Is = 2
                            amarelo = True
                            'result2.Text = "ver se há desautorização [Marca -> " & LCase(a2array(1)) & "]"
                            result2.Text = "ver se há desautorização [Marca -> genérico]"
                            result2.BackColor = Color.Yellow
                            '  labellab2.Text = "h?"
                            '  labellab2.BackColor = Color.BlueViolet
                        Case Is = 3
                            amarelo = True
                            'result3.Text = "ver se há desautorização [Marca -> " & LCase(a3array(1)) & "]"
                            result3.Text = "ver se há desautorização [Marca -> genérico]"
                            result3.BackColor = Color.Yellow
                            '  labellab3.Text = "h?"
                            '  labellab3.BackColor = Color.BlueViolet
                        Case Is = 4
                            amarelo = True
                            'result4.Text = "ver se há desautorização [Marca -> " & LCase(a4array(1)) & "]"
                            result4.Text = "ver se há desautorização [Marca -> genérico]"
                            result4.BackColor = Color.Yellow
                            ' labellab4.Text = "h?"
                            ' labellab4.BackColor = Color.BlueViolet
                            'Case Is = 5
                            '    amarelo = True
                            'result5.Text = "ver se há desautorização [Marca -> " & LCase(a5array(1)) & "]"
                            '     result5.Text = "ver se há desautorização [Marca -> genérico]"
                            '     result5.BackColor = Color.Yellow
                            '      '      labellab5.Text = "h?"
                            '    labellab5.BackColor = Color.BlueViolet
                            '  Case Is = 6
                            '   amarelo = True
                            'result6.Text = "ver se há desautorização [Marca -> " & LCase(a6array(1)) & "]"
                            ' result6.Text = "ver se há desautorização [Marca -> genérico]"
                            '  result6.BackColor = Color.Yellow
                            'labellab6.Text = "h?"
                            'labellab6.BackColor = Color.BlueViolet
                    End Select
            End Select
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("sub marca2gen: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub verifQuant(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "verificar quantidade"
4:              result1.BackColor = Color.Gray
5:          Case 2
6:              result2.Text = "verificar quantidade"
7:              result2.BackColor = Color.Gray
8:          Case 3
9:              result3.Text = "verificar quantidade"
10:             result3.BackColor = Color.Gray
11:         Case 4
12:             result4.Text = "verificar quantidade"
13:             result4.BackColor = Color.Gray
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("sub verifQuant: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub nComp(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
        vermelho = True
1:      Select Case qual
            Case 1
3:              '        labelf1.Text = "F)"
4:              '        labelf1.BackColor = Color.Red
5:          Case 2
6:              '       labelf2.Text = "F)"
7:              '      labelf2.BackColor = Color.Red
8:          Case 3
9:              '     labelf3.Text = "F)"
10:             '    labelf3.BackColor = Color.Red
11:         Case 4
12:             '   labelf4.Text = "F)"
13:             '  labelf4.BackColor = Color.Red
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("sub nComp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub aviadoNexist(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              If aviam1.Text <> "0" And aviam1.Text <> "C" And aviam1.Text <> "C0" Then
4:                  result1.BackColor = Color.Gray
5:                  result1.Text = "o código aviado não existe"
6:              End If
                If aviam1.Text = "0" Then
                    result1.BackColor = SystemColors.GradientInactiveCaption
                    result1.Text = "não aviado"
                End If
7:          Case 2
8:              If aviam2.Text <> "0" And aviam2.Text <> "C" And aviam2.Text <> "C0" Then
9:                  result2.BackColor = Color.Gray
10:                 result2.Text = "o código aviado não existe"
11:             End If
                If aviam2.Text = "0" Then
                    result2.BackColor = SystemColors.GradientInactiveCaption
                    result2.Text = "não aviado"
                End If
12:         Case 3
13:             If aviam3.Text <> "0" And aviam3.Text <> "C" And aviam3.Text <> "C0" Then
14:                 result3.BackColor = Color.Gray
15:                 result3.Text = "o código aviado não existe"
16:             End If
                If aviam3.Text = "0" Then
                    result3.BackColor = SystemColors.GradientInactiveCaption
                    result3.Text = "não aviado"
                End If
17:         Case 4
18:             If aviam4.Text <> "0" And aviam4.Text <> "C" And aviam4.Text <> "C0" Then
19:                 result4.BackColor = Color.Gray
20:                 result4.Text = "o código aviado não existe"
21:             End If
                If aviam4.Text = "0" Then
                    result4.BackColor = SystemColors.GradientInactiveCaption
                    result4.Text = "não aviado"
                End If
            Case 7
                If CodeEC1.Text <> "0" And CodeEC1.Text <> "C" And CodeEC1.Text <> "C0" Then
                    labelmedports1.Width = 200
                    labelmedports1.Location = New Point(450, 672)
                    labelmedports1.BackColor = Color.Gray
                    labelmedports1.Text = "o código aviado não existe"
                End If
22:         Case 8
23:             If CodeEC1.Text <> "0" And CodeEC1.Text <> "C" And CodeEC1.Text <> "C0" Then
                    labelmedports1.Width = 200
                    labelmedports1.Location = New Point(450, 672)
24:                 labelmedports1.BackColor = Color.Gray
25:                 labelmedports1.Text = "o código aviado não existe"
26:             End If
            Case 9
                If codEC2.Text <> "0" And codEC2.Text <> "C" And codEC2.Text <> "C0" Then
                    labelmedports.BackColor = Color.Gray
                    labelmedports.Text = "o código aviado não existe"
                End If
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub aviadoNexist: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub prescNexist(ByVal qual As Short)
        On Error GoTo MOSTRARERRO

1:      Select Case qual
            Case 1
                If P = 1 Then
3:                  If presc1.Text <> "0" Then
4:                      presc1.BackColor = Color.Red
                        result1.Text = "código prescrito desconhecido"
                        result1.BackColor = Color.Gray
                        If A = 4 Then
                            result4.Text = "código prescrito desconhecido"
                            result4.BackColor = Color.Gray
                        End If
                        If A >= 3 Then
                            result3.Text = "código prescrito desconhecido"
                            result3.BackColor = Color.Gray
                        End If
                        If A >= 2 Then
                            result2.Text = "código prescrito desconhecido"
                            result2.BackColor = Color.Gray
                        End If
5:                  End If
                End If
6:          Case 2
7:              If presc2.Text <> "0" Then
8:                  presc2.BackColor = Color.Red
9:              End If
10:         Case 3
11:             If presc3.Text <> "0" Then
12:                 presc3.BackColor = Color.Red
13:             End If
14:         Case 4
15:             If presc4.Text <> "0" Then
16:                 presc4.BackColor = Color.Red
17:             End If
18:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub prescNexist: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub msgNexist()
        On Error GoTo MOSTRARERRO
1:      MsgBox("código(s) do(s) aviado(s) não existe(m)" & vbCr & "retirar o(s) que não existe(m) e comparar de novo", MsgBoxStyle.OkOnly)
2:      result1.Text = ""
3:      result2.Text = ""
4:      result3.Text = ""
5:      result4.Text = ""
6:      Exit Sub
MOSTRARERRO:
        MsgBox("Sub msgNexist: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub fazer99(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
1:      qual.nivel = 99
2:      qual.resultado = 99
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub fazer99: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







    Sub agrupar()
1:      On Error GoTo MOSTRARERRO
        Exit Sub '20131123 para não dar erros e dar vermelho em todos os prescritos mas não vai acusar desconhecidos
2:      a1p1.grupo = 1
3:      a2p1.grupo = 1
4:      a3p1.grupo = 1
5:      a4p1.grupo = 1
6:      grupoP1 = 1
7:      If Not IsNothing(p1row) Then
8:          grupoP1dci = p1array(1).ToString
9:      End If
10:     grupoA1 = 1
11:     If Not IsNothing(a1row) Then
12:         grupoA1dci = a1row(1).ToString
13:     End If
14:     If IsNothing(p1row) Then
15:         presc1.BackColor = Color.Red
16:     End If
17:     If P >= 2 Then
18:         If Not IsNothing(p2row) Then
19:             Dim p2array = p2row.ItemArray
20:             If Not IsNothing(p1row) Then
21:                 Dim p1array = p1row.ItemArray
22:                 Select Case LCase(p2array(1).ToString)
                        Case LCase(p1array(1).ToString)
24:                         a1p2.grupo = 1
25:                         a2p2.grupo = 1
26:                         a3p2.grupo = 1
27:                         a4p2.grupo = 1
28:                         grupoP1 = grupoP1 + 1
29:                         grupoP2dci = LCase(p2array(1).ToString)
30:                         presc1.BackColor = Color.Yellow
31:                         presc2.BackColor = Color.Yellow
32:                     Case Else
33:                         a1p2.grupo = 2
34:                         a2p2.grupo = 2
35:                         a3p2.grupo = 2
36:                         a4p2.grupo = 2
37:                         grupoP2 = grupoP2 + 1
38:                         grupoP2dci = LCase(p2array(1).ToString)
39:                         End Select
40:             Else
41:                 prescNexist(1)
42:             End If
43:         Else
44:             prescNexist(2)
45:         End If
46:
47:         If P >= 3 Then
48:             If Not IsNothing(p3row) Then
49:                 Dim p3array = p3row.ItemArray
50:                 If Not IsNothing(p1row) Then
51:                     Dim p1array = p1row.ItemArray
52:                     Select Case LCase(p3array(1).ToString)
                            Case LCase(p1array(1).ToString)
54:                             a1p3.grupo = 1
55:                             a2p3.grupo = 1
56:                             a3p3.grupo = 1
57:                             a4p3.grupo = 1
58:                             grupoP1 = grupoP1 + 1
59:                             grupoP1dci = LCase(p1array(1).ToString)
60:                             presc1.BackColor = Color.Yellow
61:                             presc3.BackColor = Color.Yellow
62:                         Case Else
63:                             If Not IsNothing(p2row) Then
64:                                 Dim p2array = p2row.ItemArray
65:                                 If LCase(p3array(1).ToString) = LCase(p2array(1).ToString) Then
66:                                     a1p3.grupo = 2
67:                                     a2p3.grupo = 2
68:                                     a3p3.grupo = 2
69:                                     a4p3.grupo = 2
70:                                     grupoP2 = grupoP2 + 1
71:                                     grupoP2dci = LCase(p2array(1).ToString)
72:                                     presc3.BackColor = Color.Yellow
73:                                     presc2.BackColor = Color.Yellow
74:                                 Else
75:                                     a1p3.grupo = 3
76:                                     a2p3.grupo = 3
77:                                     a3p3.grupo = 3
78:                                     a4p3.grupo = 3
79:                                     grupoP3 = grupoP3 + 1
80:                                     grupoP3dci = LCase(p3array(1).ToString)
81:                                 End If
82:                             Else
83:                                 prescNexist(2)
84:                             End If
85:                             End Select
86:                 Else
87:                     prescNexist(1)
88:                 End If
89:             Else
90:                 prescNexist(3)
91:             End If
92:         End If
93:         If P >= 4 Then
94:             If Not IsNothing(p4row) Then
95:                 Dim p4array = p4row.ItemArray
96:                 If Not IsNothing(p1row) Then
97:                     Dim p1array = p1row.ItemArray
98:                     Select Case LCase(p4array(1).ToString)
                            Case LCase(p1array(1).ToString)
100:                            a1p4.grupo = 1
101:                            a2p4.grupo = 1
102:                            a3p4.grupo = 1
103:                            a4p4.grupo = 1
104:                            grupoP1 = grupoP1 + 1
105:                            grupoP1dci = LCase(p1array(1).ToString)
106:                            presc1.BackColor = Color.Yellow
107:                            presc4.BackColor = Color.Yellow
108:                        Case Else
109:                            If Not IsNothing(p2row) Then
110:                                Dim p2array = p2row.ItemArray
111:                                If LCase(p4array(1).ToString) = LCase(p2array(1).ToString) Then
112:                                    a1p4.grupo = 2
113:                                    a2p4.grupo = 2
114:                                    a3p4.grupo = 2
115:                                    a4p4.grupo = 2
116:                                    grupoP2 = grupoP2 + 1
117:                                    grupoP2dci = LCase(p2array(1).ToString)
118:                                    presc4.BackColor = Color.Yellow
119:                                    presc2.BackColor = Color.Yellow
120:                                Else
121:                                    If Not IsNothing(p3row) Then
122:                                        Dim p3array = p3row.ItemArray
123:                                        If LCase(p4array(1).ToString) = LCase(p3array(1).ToString) Then
124:                                            a1p4.grupo = 3
125:                                            a2p4.grupo = 3
126:                                            a3p4.grupo = 3
127:                                            a4p4.grupo = 3
128:                                            grupoP3 = grupoP3 + 1
129:                                            grupoP3dci = LCase(p3array(1).ToString)
130:                                            presc4.BackColor = Color.Yellow
131:                                            presc3.BackColor = Color.Yellow
132:                                        Else
133:                                            a1p4.grupo = 4
134:                                            a2p4.grupo = 4
135:                                            a3p4.grupo = 4
136:                                            a4p4.grupo = 4
137:                                            grupoP4 = grupoP4 + 1
138:                                            grupoP4dci = LCase(p4array(1).ToString)
139:                                        End If
140:                                    Else
141:                                        prescNexist(3)
142:                                    End If
143:                                End If
144:                            Else
145:                                prescNexist(2)
146:                            End If
147:                            End Select
148:                Else
149:                    prescNexist(1)
150:                End If
151:            Else
152:                prescNexist(4)
153:            End If
154:        End If
155:    End If
156:    Exit Sub
MOSTRARERRO:
        MsgBox("sub agrupar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Dim linha1(0 To 8)
    Dim linha2(0 To 8)
    Dim linha3(0 To 8)
    Dim linha4(0 To 8)

    'a1array.Add(linha1) - colocar isto no if semcod1.checked=true
    'dá erro. corrigir e transpor para os outros aviam's e os outros subaviam's
    'em todas as verificações de prescX.text e de pXrow(0) e pXarray(0) tem de se colocar antes o if semcodX.checked=True

































































    Sub IndicarComoMostrado(ByVal qualP As Short, ByVal qualA As Short)
1:      On Error GoTo MOSTRARERRO
2:      Select Case qualP
            Case 1
3:              Select Case qualA
                    Case 1
4:                      a1p1.mostrado = True
5:                      a1p2.mostrado = True
6:                      a1p3.mostrado = True
7:                      a1p4.mostrado = True
8:                      'a2p1.mostrado = True
9:                      'a3p1.mostrado = True
10:                     'a4p1.mostrado = True
                    Case 2
11:                     a2p1.mostrado = True
12:                     a2p2.mostrado = True
13:                     a2p3.mostrado = True
14:                     a2p4.mostrado = True
15:                     'a1p1.mostrado = True
16:                     'a3p1.mostrado = True
17:                     'a4p1.mostrado = True
                    Case 3
19:                     a3p1.mostrado = True
20:                     a3p2.mostrado = True
21:                     a3p3.mostrado = True
22:                     a3p4.mostrado = True
23:                     'a2p1.mostrado = True
24:                     'a1p1.mostrado = True
25:                     'a4p1.mostrado = True
                    Case 4
26:                     a4p1.mostrado = True
27:                     a4p2.mostrado = True
28:                     a4p3.mostrado = True
29:                     a4p4.mostrado = True
30:                     'a2p1.mostrado = True
31:                     'a3p1.mostrado = True
32:                     'a1p1.mostrado = True
33:                     End Select
            Case 2
34:             Select Case qualA
                    Case 1
35:                     a1p1.mostrado = True
36:                     a1p2.mostrado = True
37:                     a1p3.mostrado = True
38:                     a1p4.mostrado = True
39:                     'a2p2.mostrado = True
40:                     'a3p2.mostrado = True
41:                     'a4p2.mostrado = True
                    Case 2
42:                     a2p1.mostrado = True
43:                     a2p2.mostrado = True
44:                     a2p3.mostrado = True
45:                     a2p4.mostrado = True
46:                     'a1p2.mostrado = True
47:                     'a3p2.mostrado = True
48:                     'a4p2.mostrado = True
49:                 Case 3
50:                     a3p1.mostrado = True
51:                     a3p2.mostrado = True
52:                     a3p3.mostrado = True
53:                     a3p4.mostrado = True
54:                     'a2p2.mostrado = True
55:                     'a1p2.mostrado = True
56:                     'a4p2.mostrado = True
                    Case 4
57:                     a4p1.mostrado = True
58:                     a4p2.mostrado = True
59:                     a4p3.mostrado = True
60:                     a4p4.mostrado = True
61:                     'a2p2.mostrado = True
62:                     'a3p2.mostrado = True
63:                     'a1p2.mostrado = True
64:                     End Select
            Case 3
65:             Select Case qualA
                    Case 1
66:                     a1p1.mostrado = True
67:                     a1p2.mostrado = True
68:                     a1p3.mostrado = True
69:                     a1p4.mostrado = True
70:                     'a2p3.mostrado = True
71:                     'a3p3.mostrado = True
72:                     'a4p3.mostrado = True
                    Case 2
73:                     a2p1.mostrado = True
74:                     a2p2.mostrado = True
75:                     a2p3.mostrado = True
76:                     a2p4.mostrado = True
77:                     'a1p3.mostrado = True
78:                     'a3p3.mostrado = True
79:                     'a4p3.mostrado = True
                    Case 3
80:                     a3p1.mostrado = True
81:                     a3p2.mostrado = True
82:                     a3p3.mostrado = True
83:                     a3p4.mostrado = True
84:                     'a2p3.mostrado = True
85:                     'a1p3.mostrado = True
86:                     'a4p3.mostrado = True
                    Case 4
87:                     a4p1.mostrado = True
88:                     a4p2.mostrado = True
89:                     a4p3.mostrado = True
90:                     a4p4.mostrado = True
91:                     'a2p3.mostrado = True
92:                     'a3p3.mostrado = True
93:                     'a1p3.mostrado = True
                End Select
            Case 4
94:             Select Case qualA
                    Case 1
95:                     a1p1.mostrado = True
96:                     a1p2.mostrado = True
97:                     a1p3.mostrado = True
98:                     a1p4.mostrado = True
99:                     'a2p4.mostrado = True
100:                    'a3p4.mostrado = True
101:                    'a4p4.mostrado = True
                    Case 2
102:                    a2p1.mostrado = True
103:                    a2p2.mostrado = True
104:                    a2p3.mostrado = True
105:                    a2p4.mostrado = True
106:                    'a1p4.mostrado = True
107:                    'a3p4.mostrado = True
108:                    'a4p4.mostrado = True
                    Case 3
109:                    a3p1.mostrado = True
110:                    a3p2.mostrado = True
111:                    a3p3.mostrado = True
112:                    a3p4.mostrado = True
113:                    'a2p4.mostrado = True
114:                    'a1p4.mostrado = True
115:                    'a4p4.mostrado = True
                    Case 4
116:                    a4p1.mostrado = True
117:                    a4p2.mostrado = True
118:                    a4p3.mostrado = True
119:                    a4p4.mostrado = True
120:                    'a2p4.mostrado = True
121:                    'a3p4.mostrado = True
122:                    'a1p4.mostrado = True
123:                    End Select
124:            End Select
125:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub IndicarComoMostrado: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub ApagarRepetido(ByVal qualP As Short, ByVal qualA As Short)
        On Error GoTo MOSTRARERRO
        Select Case qualP
            Case 1
                Select Case qualA
                    Case 1
                        cruz(1, 1)
                        av1.nivel = a1p1.nivel
                        If Not a1p1.mostrado = True Then
                            If a2p1.nivel <= a1p1.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(1, 2)
                            End If
                            If a3p1.nivel <= a1p1.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(1, 3)
                            End If
                            If a4p1.nivel <= a1p1.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(1, 4)
                            End If
                            IndicarComoMostrado(1, 1)
                        End If
                    Case 2
                        cruz(1, 2)
                        av2.nivel = a2p1.nivel
                        If Not a2p1.mostrado = True Then
                            If a1p1.nivel <= a2p1.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(1, 1)
                            End If
                            If a3p1.nivel <= a2p1.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(1, 3)
                            End If
                            If a4p1.nivel <= a2p1.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(1, 4)
                            End If
                            IndicarComoMostrado(1, 2)
                        End If
                    Case 3
                        cruz(1, 3)
                        av3.nivel = a3p1.nivel
                        If Not a3p1.mostrado = True Then
                            If a1p1.nivel <= a3p1.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(1, 1)
                            End If
                            If a2p1.nivel <= a3p1.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(1, 2)
                            End If
                            If a4p1.nivel <= a3p1.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(1, 4)
                            End If
                            IndicarComoMostrado(1, 3)
                        End If
                    Case 4
                        cruz(1, 4)
                        av4.nivel = a4p1.nivel
                        If Not a4p1.mostrado = True Then
                            If a1p1.nivel <= a4p1.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(1, 1)
                            End If
                            If a2p1.nivel <= a4p1.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(1, 2)
                            End If
                            If a3p1.nivel <= a4p1.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(1, 3)
                            End If
                            IndicarComoMostrado(1, 4)
                        End If
                End Select
            Case 2
                Select Case qualA
                    Case 1
                        cruz(2, 1)
                        av1.nivel = a1p2.nivel
                        If Not a1p2.mostrado = True Then
                            If a2p2.nivel <= a1p2.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(2, 2)
                            End If
                            If a3p2.nivel <= a1p2.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(2, 3)
                            End If
                            If a4p2.nivel <= a1p2.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(2, 4)
                            End If
                            IndicarComoMostrado(2, 1)
                        End If
                    Case 2
                        cruz(2, 2)
                        av2.nivel = a2p2.nivel
                        If Not a2p2.mostrado = True Then
                            If a1p2.nivel <= a2p2.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(2, 1)
                            End If
                            If a3p2.nivel <= a2p2.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(2, 3)
                            End If
                            If a4p2.nivel <= a2p2.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(2, 4)
                            End If
                            IndicarComoMostrado(2, 2)
                        End If
                    Case 3
                        cruz(2, 3)
                        av3.nivel = a3p2.nivel
                        If Not a3p2.mostrado = True Then
                            If a1p2.nivel <= a3p2.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(2, 1)
                            End If
                            If a2p2.nivel <= a3p2.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(2, 2)
                            End If
                            If a4p2.nivel <= a3p2.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(2, 4)
                            End If
                            IndicarComoMostrado(2, 3)
                        End If
                    Case 4
                        cruz(2, 4)
                        av4.nivel = a4p2.nivel
                        If Not a4p2.mostrado = True Then
                            If a1p2.nivel <= a4p2.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(2, 1)
                            End If
                            If a2p2.nivel <= a4p2.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(2, 2)
                            End If
                            If a3p2.nivel <= a4p2.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(2, 3)
                            End If
                            IndicarComoMostrado(2, 4)
                        End If
                End Select
            Case 3
                Select Case qualA
                    Case 1
                        cruz(3, 1)
                        av1.nivel = a1p3.nivel
                        If Not a1p3.mostrado = True Then
                            If a2p3.nivel <= a1p3.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(3, 2)
                            End If
                            If a3p3.nivel <= a1p3.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(3, 3)
                            End If
                            If a4p3.nivel <= a1p3.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(3, 4)
                            End If
                            IndicarComoMostrado(3, 1)
                        End If
                    Case 2
                        cruz(3, 2)
                        av2.nivel = a2p3.nivel
                        If Not a2p3.mostrado = True Then
                            If a1p3.nivel <= a2p3.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(3, 1)
                            End If
                            If a3p3.nivel <= a2p3.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(3, 3)
                            End If
                            If a4p3.nivel <= a2p3.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(3, 4)
                            End If
                            IndicarComoMostrado(3, 2)
                        End If
                    Case 3
                        cruz(3, 3)
                        av3.nivel = a3p3.nivel
                        If Not a3p3.mostrado = True Then
                            If a1p3.nivel <= a3p3.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(3, 1)
                            End If
                            If a2p3.nivel <= a3p3.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(3, 2)
                            End If
                            If a4p3.nivel <= a3p3.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(3, 4)
                            End If
                            IndicarComoMostrado(3, 3)
                        End If
                    Case 4
                        cruz(3, 4)
                        av4.nivel = a4p3.nivel
                        If Not a4p3.mostrado = True Then
                            If a1p3.nivel <= a4p3.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(3, 1)
                            End If
                            If a2p3.nivel <= a4p3.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(3, 2)
                            End If
                            If a3p3.nivel <= a4p3.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(3, 3)
                            End If
                            IndicarComoMostrado(3, 4)
                        End If
                End Select
            Case 4
                Select Case qualA
                    Case 1
                        cruz(4, 1)
                        av1.nivel = a1p4.nivel
                        If Not a1p4.mostrado = True Then
                            If a2p4.nivel <= a1p4.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(4, 2)
                            End If
                            If a3p4.nivel <= a1p4.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(4, 3)
                            End If
                            If a4p4.nivel <= a1p4.nivel Then
                                dciDif(1)
                                'IndicarComoMostrado(4, 4)
                            End If
                            IndicarComoMostrado(4, 1)
                        End If
                    Case 2
                        cruz(4, 2)
                        av2.nivel = a2p4.nivel
                        If Not a2p4.mostrado = True Then
                            If a1p4.nivel <= a2p4.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(4, 1)
                            End If
                            If a3p4.nivel <= a2p4.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(4, 3)
                            End If
                            If a4p4.nivel <= a2p4.nivel Then
                                dciDif(2)
                                'IndicarComoMostrado(4, 4)
                            End If
                            IndicarComoMostrado(4, 2)
                        End If
                    Case 3
                        cruz(4, 3)
                        av3.nivel = a3p4.nivel
                        If Not a3p4.mostrado = True Then
                            If a1p4.nivel <= a3p4.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(4, 1)
                            End If
                            If a2p4.nivel <= a3p4.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(4, 2)
                            End If
                            If a4p4.nivel <= a3p4.nivel Then
                                dciDif(3)
                                'IndicarComoMostrado(4, 4)
                            End If
                            IndicarComoMostrado(4, 3)
                        End If
                    Case 4
                        cruz(4, 4)
                        av4.nivel = a4p4.nivel
                        If Not a4p4.mostrado = True Then
                            If a1p4.nivel <= a4p4.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(4, 1)
                            End If
                            If a2p4.nivel <= a4p4.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(4, 2)
                            End If
                            If a3p4.nivel <= a4p4.nivel Then
                                dciDif(4)
                                'IndicarComoMostrado(4, 3)
                            End If
                            IndicarComoMostrado(4, 4)
                        End If
                End Select
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub ApagraRepetido: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Dim paramiloidose As Boolean

    'usados quando não tinha organismos, para permitir comparticipar tudo na paramiloidose com paramiloidoseRB
    Private Sub ParamiloidoseRB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      If but42.Checked = True Then
2:          paramiloidose = True
3:      End If
4:      If but42.Checked = False Then
6:          paramiloidose = False
7:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    '    Sub alzheimer(ByVal qual As Object)
    '        On Error GoTo MOSTRARERRO
    '    'qual entra no formato aXrow(0) ou aXarray(0)...
    '1:      MsgBox(qual & " necessita de" & vbCr _
    '               & "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
    '               & "e especialidade de neurologia ou psiquiatria" & vbCr _
    '               & "para ser comparticipado")
    '2:      Select Case qual
    '            Case 1
    '3:              aviam1.BackColor = Color.Purple
    '4:          Case 2
    '5:              aviam2.BackColor = Color.Purple
    '6:          Case 3
    '7:              aviam3.BackColor = Color.Purple
    '8:          Case 4
    '9:              aviam4.BackColor = Color.Purple
    '10:             End Select
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub


    '  Sub gastro(ByVal qual As Object)
    '      On Error GoTo MOSTRARERRO
    '  'qual entra no formato aXrow(0) ou aXarray(0)...
    '1:      MsgBox(qual & " necessita de especialidade" & vbCr _
    '               & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
    '               & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
    '2:      Select Case qual
    '            Case 1
    '3:              aviam1.BackColor = Color.Purple
    '4:          Case 2
    '5:              aviam2.BackColor = Color.Purple
    '6:          Case 3
    '7:              aviam3.BackColor = Color.Purple
    '8:          Case 4
    '9:              aviam4.BackColor = Color.Purple
    '10:             End Select
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    'Sub bipolar(ByVal qual As Object)
    '   On Error GoTo MOSTRARERRO
    'qual entra no formato aXrow(0) ou aXarray(0)...
    '1:      MsgBox(qual & " necessita de especialidade" & vbCr _
    '               & "de psiquiatria ou de neurologia" & vbCr _
    '               & "no caso de comparticipado com o despacho nº. 21094/99")
    '2:      Select Case qual
    '            Case 1
    '3:              aviam1.BackColor = Color.Purple
    '4:          Case 2
    '5:              aviam2.BackColor = Color.Purple
    '6:          Case 3
    '7:              aviam3.BackColor = Color.Purple
    '8:          Case 4
    '9:              aviam4.BackColor = Color.Purple
    '10:             End Select
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub



    'Sub espondilite(ByVal qual As Object)
    '   On Error GoTo MOSTRARERRO
    'qual entra no formato aXrow(0) ou aXarray(0)...
    '1:      MsgBox(qual & " necessita de especialidade" & vbCr _
    '             & "de reumatologia ou de medicina interna" & vbCr _
    '             & "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
    '4:      Select Case qual
    '            Case 1
    '6:              aviam1.BackColor = Color.Purple
    '7:          Case 2
    '8:              aviam2.BackColor = Color.Purple
    '9:          Case 3
    '10:             aviam3.BackColor = Color.Purple
    '11:         Case 4
    '12:             aviam4.BackColor = Color.Purple
    '13:             End Select
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub


    'Sub portaria(ByVal qual As Object, ByVal desp As String)
    '  On Error GoTo MOSTRARERRO
    ' MsgBox(qual & " pode levar o despacho nº." & desp)
    'Select Case qual
    '   Case 1
    '      aviam1.BackColor = Color.Beige
    ' Case 2
    '    aviam2.BackColor = Color.Beige
    ' Case 3
    '     aviam3.BackColor = Color.Beige
    ' Case 4
    '     aviam4.BackColor = Color.Beige
    '     End Select
    'Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub





    Sub AvisarDespachos()
        On Error GoTo MOSTRARERRO
        If P = 1 Then
            If A >= 1 Then
                If Not IsNothing(a1row) Then
                    'If a1row(11) = True Or a1row(13) = True Or a1row(14) = True Or a1row(15) = True Or a1row(16) = True Or a1row(17) = True Then
                    If a1row(10) = True Then
                        If Not a1_4250 = True And av1.nivel <= 3 Then
                            portariado(1, 4250)
                            a1_4250 = True
                        End If
                    End If
                    If a1row(11) = True Then
                        If Not a1_1234 = True And av1.nivel <= 3 Then
                            portariado(1, 1234)
                            a1_1234 = True
                        End If
                    End If
                    If a1row(16) = True Then
                        If Not a1_14123 = True And av1.nivel <= 3 Then
                            portariado(1, 14123)
                            a1_14123 = True
                        End If
                    End If
                    ' If a1row(17) = True Then
                    ' If port1474 = True Then
                    ' If Not a1_1474ad = True And av1.nivel <= 3 Then
                    'a1_1474ad = True
                    '      End If
                    ' End If
                    'End If
                    'If a1row(18) = True Then
                    ' If port1474 = True Then
                    ' If Not a1_1474nl = True And av1.nivel <= 3 Then
                    'a1_1474nl = True
                    ''           End If
                    'End If
                    'End If
                    If a1row(13) = True Then
                        If Not a1_10279 = True And av1.nivel <= 3 Then
                            portariado(1, 10279)
                            a1_10279 = True
                        End If
                    End If
                    If a1row(14) = True Then
                        If Not a1_10279 = True And av1.nivel <= 3 Then
                            portariado(1, 10279)
                            a1_10279 = True
                        End If
                    End If
                    If a1row(15) = True Then
                        If Not a1_10910 = True And av1.nivel <= 3 Then
                            portariado(1, 10910)
                            a1_10910 = True
                        End If
                    End If
                    If a1row(12) = True Then
                        If Not a1_21094 = True And av1.nivel <= 3 Then
                            portariado(1, 21094)
                            a1_21094 = True
                        End If
                    End If
                End If
            End If
            If A >= 2 And Not IsNothing(a2row) Then
                'If a2row(11) = True Or a2row(13) = True Or a2row(14) = True Or a2row(15) = True Or a2row(16) = True Or a2row(17) = True Then
                If a2row(10) = True Then
                    If Not a2_4250 = True And av2.nivel <= 3 Then
                        portariado(2, 4250)
                        a2_4250 = True
                    End If
                End If
                If a2row(11) = True Then
                    If Not a2_1234 = True And av2.nivel <= 3 Then
                        portariado(2, 1234)
                        a2_1234 = True
                    End If
                End If
                If a2row(16) = True Then
                    If Not a2_14123 = True And av2.nivel <= 3 Then
                        portariado(2, 14123)
                        a2_14123 = True
                    End If
                End If
                ' If a2row(17) = True Then
                ' If port1474 = True Then
                'If Not a2_1474ad = True And av2.nivel <= 3 Then
                'a2_1474ad = True
                '   End If
                'End If
                'End If
                'If a2row(18) = True Then
                ' If port1474 = True Then
                ' If Not a2_1474nl = True And av2.nivel <= 3 Then
                'a2_1474nl = True
                '       End If
                ' End If
                ' End If
                If a2row(13) = True Then
                    If Not a2_10279 = True And av2.nivel <= 3 Then
                        portariado(2, 10279)
                        a2_10279 = True
                    End If
                End If
                If a2row(14) = True Then
                    If Not a2_10279 = True And av2.nivel <= 3 Then
                        portariado(2, 10279)
                        a2_10279 = True
                    End If
                End If
                If a2row(15) = True Then
                    If Not a2_10910 = True And av2.nivel <= 3 Then
                        portariado(2, 10910)
                        a2_10910 = True
                    End If
                End If
                If a2row(12) = True Then
                    If Not a2_21094 = True And av2.nivel <= 3 Then
                        portariado(2, 21094)
                        a2_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 3 And Not IsNothing(a3row) Then
                'If a3row(11) = True Or a3row(13) = True Or a3row(14) = True Or a3row(15) = True Or a3row(16) = True Or a3row(17) = True Then
                If a3row(10) = True Then
                    If Not a3_4250 = True And av3.nivel <= 3 Then
                        portariado(3, 4250)
                        a3_4250 = True
                    End If
                End If
                If a3row(11) = True And av3.nivel <= 3 Then
                    If Not a3_1234 = True Then
                        portariado(3, 1234)
                        a3_1234 = True
                    End If
                End If
                If a3row(16) = True And av3.nivel <= 3 Then
                    If Not a3_14123 = True Then
                        portariado(3, 14123)
                        a3_14123 = True
                    End If
                End If
                '         If a3row(17) = True Then
                ' If port1474 = True Then
                ' If Not a3_1474ad = True And av3.nivel <= 3 Then
                'a3_1474ad = True
                '       End If
                'End If
                'End If
                'If a3row(18) = True Then
                ' If port1474 = True Then
                ' If Not a3_1474nl = True And av3.nivel <= 3 Then
                'a3_1474nl = True
                '    End If
                'End If
                '   End If
                If a3row(13) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        portariado(3, 10279)
                        a3_10279 = True
                    End If
                End If
                If a3row(14) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        portariado(3, 10279)
                        a3_10279 = True
                    End If
                End If
                If a3row(15) = True And av3.nivel <= 3 Then
                    If Not a3_10910 = True Then
                        portariado(3, 10910)
                        a3_10910 = True
                    End If
                End If
                If a3row(12) = True Then
                    If Not a3_21094 = True And av3.nivel <= 3 Then
                        portariado(3, 21094)
                        a3_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 4 And Not IsNothing(a4row) Then
                'If a4row(11) = True Or a4row(13) = True Or a4row(14) = True Or a4row(15) = True Or a4row(16) = True Or a4row(17) = True Then
                If a4row(10) = True Then
                    If Not a4_4250 = True And av4.nivel <= 3 Then
                        portariado(4, 4250)
                        a4_4250 = True
                    End If
                End If
                If a4row(11) = True And av4.nivel <= 3 Then
                    If Not a4_1234 = True Then
                        portariado(4, 1234)
                        a4_1234 = True
                    End If
                End If
                If a4row(16) = True And av4.nivel <= 3 Then
                    If Not a4_14123 = True Then
                        portariado(4, 14123)
                        a4_14123 = True
                    End If
                End If
                '             If a4row(17) = True Then
                ' If port1474 = True Then
                ' If Not a4_1474ad = True And av4.nivel <= 3 Then
                ' a4_1474ad = True
                '       End If
                'End If
                ' End If
                ' If a4row(18) = True Then
                'If port1474 = True Then
                ' If Not a4_1474nl = True And av4.nivel <= 3 Then
                'a4_1474nl = True
                '   End If
                'End If
                ' End If
                If a4row(13) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        portariado(4, 10279)
                        a4_10279 = True
                    End If
                End If
                If a4row(14) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        portariado(4, 10279)
                        a4_10279 = True
                    End If
                End If
                If a4row(15) = True And av4.nivel <= 3 Then
                    If Not a4_10910 = True Then
                        portariado(4, 10910)
                        a4_10910 = True
                    End If
                End If
                If a4row(12) = True Then
                    If Not a4_21094 = True And av4.nivel <= 3 Then
                        portariado(4, 21094)
                        a4_21094 = True
                    End If
                End If
            End If
        Else
            If A >= 1 Then
                If Not IsNothing(a1row) Then
                    'If a1row(11) = True Or a1row(13) = True Or a1row(14) = True Or a1row(15) = True Or a1row(16) = True Or a1row(17) = True Then
                    If a1row(10) = True Then
                        If Not a1_4250 = True And av1.nivel <= 3 Then
                            portariado(1, 4250)
                            a1_4250 = True
                        End If
                    End If
                    If a1row(11) = True And av1.nivel <= 3 Then
                        If Not a1_1234 = True Then
                            portariado(1, 1234)
                            a1_1234 = True
                        End If
                    End If
                    If a1row(16) = True And av1.nivel <= 3 Then
                        If Not a1_14123 = True Then
                            portariado(1, 14123)
                            a1_14123 = True
                        End If
                    End If
                    If a1row(13) = True And av1.nivel <= 3 Then
                        If Not a1_10279 = True Then
                            portariado(1, 10279)
                            a1_10279 = True
                        End If
                    End If
                    If a1row(14) = True And av1.nivel <= 3 Then
                        If Not a1_10279 = True Then
                            portariado(1, 10279)
                            a1_10279 = True
                        End If
                    End If
                    If a1row(15) = True And av1.nivel <= 3 Then
                        If Not a1_10910 = True Then
                            portariado(1, 10910)
                            a1_10910 = True
                        End If
                    End If
                    '       If a1row(17) = True And av1.nivel <= 3 Then
                    ' If Not a1_1474ad = True Then
                    ' If port1474 = True Then
                    'portariado(1, 1474)
                    'a1_1474ad = True
                    '          End If
                    'End If
                    'End If
                    'If a1row(18) = True And av1.nivel <= 3 Then
                    ' If Not a1_1474nl = True Then
                    ' If port1474 = True Then
                    'portariado(1, 1474)
                    'a1_1474nl = True
                    '             End If
                    'End If
                    ' End If
                    If a1row(12) = True And av1.nivel <= 3 Then
                        If Not a1_21094 = True Then
                            portariado(1, 21094)
                            a1_21094 = True
                        End If
                    End If
                End If
            End If
            If A >= 2 And Not IsNothing(a2row) Then
                'If a2row(11) = True Or a2row(13) = True Or a2row(14) = True Or a2row(15) = True Or a2row(16) = True Or a2row(17) = True Then
                If a2row(10) = True And av2.nivel <= 3 Then
                    If Not a2_4250 = True Then
                        portariado(2, 4250)
                        a2_4250 = True
                    End If
                End If
                If a2row(11) = True And av2.nivel <= 3 Then
                    If Not a2_1234 = True Then
                        portariado(2, 1234)
                        a2_1234 = True
                    End If
                End If
                If a2row(16) = True And av2.nivel <= 3 Then
                    If Not a2_14123 = True Then
                        portariado(2, 14123)
                        a2_14123 = True
                    End If
                End If
                If a2row(13) = True And av2.nivel <= 3 Then
                    If Not a2_10279 = True Then
                        portariado(2, 10279)
                        a2_10279 = True
                    End If
                End If
                If a2row(14) = True And av2.nivel <= 3 Then
                    If Not a2_10279 = True Then
                        portariado(2, 10279)
                        a2_10279 = True
                    End If
                End If
                If a2row(15) = True And av2.nivel <= 3 Then
                    If Not a2_10910 = True Then
                        portariado(2, 10910)
                        a2_10910 = True
                    End If
                End If
                '          If a2row(17) = True And av2.nivel <= 3 Then
                '              If Not a2_1474ad = True Then
                ' If port1474 = True Then
                'portariado(2, 1474)
                'a2_1474ad = True
                ''        End If
                '    End If
                '    End If
                'If a2row(18) = True And av2.nivel <= 3 Then
                'If Not a2_1474nl = True Then
                'If port1474 = True Then
                'portariado(2, 1474)
                '                a2_1474nl = True
                '           End If
                '      End If
                '     End If
                If a2row(12) = True And av2.nivel <= 3 Then
                    If Not a2_21094 = True Then
                        portariado(2, 21094)
                        a2_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 3 And Not IsNothing(a3row) Then
                'If a3row(11) = True Or a3row(13) = True Or a3row(14) = True Or a3row(15) = True Or a3row(16) = True Or a3row(17) = True Then
                If a3row(10) = True And av3.nivel <= 3 Then
                    If Not a3_4250 = True Then
                        portariado(3, 4250)
                        a3_4250 = True
                    End If
                End If
                If a3row(11) = True And av3.nivel <= 3 Then
                    If Not a3_1234 = True Then
                        portariado(3, 1234)
                        a3_1234 = True
                    End If
                End If
                If a3row(16) = True And av3.nivel <= 3 Then
                    If Not a3_14123 = True Then
                        portariado(3, 14123)
                        a3_14123 = True
                    End If
                End If
                If a3row(13) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        portariado(3, 10279)
                        a3_10279 = True
                    End If
                End If
                If a3row(14) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        portariado(3, 10279)
                        a3_10279 = True
                    End If
                End If
                If a3row(15) = True And av3.nivel <= 3 Then
                    If Not a3_10910 = True Then
                        portariado(3, 10910)
                        a3_10910 = True
                    End If
                End If
                '         If a3row(17) = True And av3.nivel <= 3 Then
                '             If Not a3_1474ad = True Then
                ' If port1474 = True Then
                ' portariado(3, 1474)
                '               a3_1474ad = True
                '           End If
                '       End If
                '       End If
                'If a3row(18) = True And av3.nivel <= 3 Then
                'If Not a3_1474nl = True Then
                ' If port1474 = True Then
                'portariado(3, 1474)
                '           a3_1474nl = True
                '       End If
                '   End If
                '   End If
                If a3row(12) = True And av3.nivel <= 3 Then
                    If Not a3_21094 = True Then
                        portariado(3, 21094)
                        a3_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 4 And Not IsNothing(a4row) Then
                'If a4row(11) = True Or a4row(13) = True Or a4row(14) = True Or a4row(15) = True Or a4row(16) = True Or a4row(17) = True Then
                If a4row(10) = True And av4.nivel <= 3 Then
                    If Not a4_4250 = True Then
                        portariado(4, 4250)
                        a4_4250 = True
                    End If
                End If
                If a4row(11) = True And av4.nivel <= 3 Then
                    If Not a4_1234 = True Then
                        portariado(4, 1234)
                        a4_1234 = True
                    End If
                End If
                If a4row(16) = True And av4.nivel <= 3 Then
                    If Not a4_14123 = True Then
                        portariado(4, 14123)
                        a4_14123 = True
                    End If
                End If
                If a4row(13) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        portariado(4, 10279)
                        a4_10279 = True
                    End If
                End If
                If a4row(14) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        portariado(4, 10279)
                        a4_10279 = True
                    End If
                End If
                If a4row(15) = True And av4.nivel <= 3 Then
                    If Not a4_10910 = True Then
                        portariado(4, 10910)
                        a4_10910 = True
                    End If
                End If
                '          If a4row(17) = True And av4.nivel <= 3 Then
                '              If Not a4_1474ad = True Then
                'If port1474 = True Then
                ' portariado(4, 1474)
                '             a4_1474ad = True
                '         End If
                '     End If
                '     End If
                ' If a4row(18) = True And av4.nivel <= 3 Then
                ' If Not a4_1474nl = True Then
                ' If port1474 = True Then
                'portariado(4, 1474)
                '   a4_1474nl = True
                ' End' If
                'End If
                'End If
                If a4row(12) = True And av4.nivel <= 3 Then
                    If Not a4_21094 = True Then
                        portariado(4, 21094)
                        a4_21094 = True
                    End If
                End If
                'End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub AvisarDespachos: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub TresPresc()
        On Error GoTo MOSTRARERRO
1:      If grupoP1 >= 3 Then
2:          'MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
3:      End If
4:      If grupoP2 >= 3 Then
5:          'MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
6:      End If
7:      If grupoP3 >= 3 Then
8:          'MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
9:      End If
10:     If grupoP4 >= 3 Then
11:         'MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
12:     End If
13:     If grupoA1 >= 3 Then
14:         'MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
15:     End If
16:     If grupoA2 >= 3 Then
17:         'MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
18:     End If
19:     If grupoA3 >= 3 Then
20:         'MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
21:     End If
22:     If grupoA4 >= 3 Then
23:         'MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
24:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub TrestPresc: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'a troca de apresentação é avaliada em função da de administração (acrescentei depois o subvia que avalia a forma mesmo).
    'cada uma tem muitas formas. aqui se associa a via à forma


    Public Function subvia(ByVal forma As String) As Short 'era via e funcionava no v3.0
        On Error GoTo mostrarerro
        Dim lforma As String
        lforma = LCase(forma)
1:      Select Case lforma
            Case "cápsula dura"
                subvia = 101
            Case "cápsula mole"
                subvia = 101
            Case "comprimido"
                subvia = 101
            Case "comprimido revestido"
                subvia = 101
            Case "comprimido revestido + comprimido"
                subvia = 101
            Case "comprimido revestido por película"
                subvia = 101
            Case "cápsula"
                subvia = 101
            Case "cápsula de libertação modificada"
                subvia = 102
            Case "cápsula de libertação prolongada"
                subvia = 102
            Case "cápsula dura de libertação prolongada"
                subvia = 102
            Case "cápsula mole de libertação modificada"
                subvia = 102
            Case "comprimido de libertação modificada"
                subvia = 102
            Case "comprimido de libertação prolongada"
                subvia = 102
            Case "comprimido de libertação prolongada revestido por película"
                subvia = 102
            Case "cápsula dura de libertação modificada"
                subvia = 102
            Case "cápsula gastrorresistente"
                subvia = 103
            Case "cápsula mole gastrorresistente"
                subvia = 103
            Case "comprimido gastrorresistente"
                subvia = 103
            Case "cápsula dura gastrorresistente"
                subvia = 103
            Case "comprimido para chupar"
                subvia = 106
            Case "comprimido para mastigar"
                subvia = 106
            Case "goma para mascar medicamentosa"
                subvia = 106
            Case "liofilizado oral"
                subvia = 106
            Case "película orodispersível"
                subvia = 106
            Case "comprimido orodispersível"
                subvia = 106
            Case "pastilha"
                subvia = 108
            Case "comprimido sublingual"
                subvia = 109
            Case "comprimido + suspensão oral"
                subvia = 110
            Case "gel oral"
                subvia = 111
            Case "comprimido dispersível"
                subvia = 112
            Case "comprimido dispersível ou para mastigar"
                subvia = 112
            Case "comprimido efervescente"
                subvia = 112
            Case "comprimido solúvel"
                subvia = 112
            Case "gotas orais, suspensão"
                subvia = 112
            Case "granulado"
                subvia = 112
            Case "granulado efervescente"
                subvia = 112
            Case "granulado para solução oral"
                subvia = 112
            Case "granulado para suspensão oral"
                subvia = 112
            Case "granulado para xarope"
                subvia = 112
            Case "pó e solvente para suspensão oral"
                subvia = 112
            Case "pó efervescente"
                subvia = 112
            Case "pó oral"
                subvia = 112
            Case "pó para solução oral"
                subvia = 112
            Case "pó para suspensão oral"
                subvia = 112
            Case "solução oral"
                subvia = 112
            Case "xarope"
                subvia = 112
            Case "gotas orais, solução"
                subvia = 112
            Case "pó e solvente para solução oral"
                subvia = 112
            Case "suspensão oral"
                subvia = 112

            Case "solução para pulverização bucal"
                subvia = 115
            Case "suspensão dental"
                subvia = 116
            Case "granulado de libertação modificada"
                subvia = 117
            Case "granulado de libertação prolongada"
                subvia = 117
            Case "pasta dentífrica"
                subvia = 118
            Case "gel dental"
                subvia = 119
            Case "champô"
                subvia = 200
            Case "espuma cutânea"
                subvia = 201
            Case "líquido cutâneo"
                subvia = 201
            Case "solução cutânea"
                subvia = 201
            Case "emulsão cutânea"
                subvia = 201
            Case "gel"
                subvia = 202
            Case "pomada"
                subvia = 202
            Case "creme"
                subvia = 202
            Case "pó para pulverização cutânea"
                subvia = 203
            Case "solução para pulverização cutânea"
                subvia = 203
            Case "pó cutâneo"
                subvia = 203
            Case "sistema transdérmico"
                subvia = 204
            Case "penso impregnado"
                subvia = 204
            Case "colírio, solução"
                subvia = 300
            Case "colírio, suspensão"
                subvia = 300
            Case "colírio, pó e solvente para solução"
                subvia = 300
            Case "colírio de libertação prolongada"
                subvia = 301
            Case "colírio de acção prolongada"
                subvia = 301
            Case "pomada oftálmica"
                subvia = 302
            Case "gel oftálmico"
                subvia = 302
            Case "gotas auriculares, suspensão"
                subvia = 400
            Case "solução para pulverização auricular"
                subvia = 400
            Case "gotas auriculares, solução"
                subvia = 400
            Case "gás para inalação"
                subvia = 500

            Case "solução para pulverização nasal"
                subvia = 501
            Case "suspensão para pulverização nasal"
                subvia = 501
            Case "gotas nasais, solução"
                subvia = 501

                'infarmed descodificou do 503
            Case "cápsula para inalação"
                subvia = 504
            Case "pó nasal"
                subvia = 503
            Case "pó para inalação"
                subvia = 503
            Case "cápsula para inalação por vaporização"
                subvia = 503
            Case "pó para inalação, cápsula"
                subvia = 503
            Case "pó para inalação, cápsula dura"
                subvia = 503
            Case "solução para inalação por nebulização"
                subvia = 503
            Case "solução para inalação por vaporização"
                subvia = 503
            Case "solução pressurizada para inalação"
                subvia = 503
            Case "suspensão"
                subvia = 503
            Case "suspensão pressurizada para inalação"
                subvia = 503
            Case "líquido para inalação por vaporização"
                subvia = 503
            Case "suspensão para inalação por nebulização"
                subvia = 503

                ' infarmed não codificou por isso criei eu - não pus no 503 para poder dar erro
            Case "pó para inalação em recipiente unidose"
                subvia = 504

            Case "comprimido vaginal"
                subvia = 600
            Case "cápsula mole vaginal"
                subvia = 600
            Case "creme vaginal"
                subvia = 602
            Case "espuma vaginal"
                subvia = 602
            Case "gel vaginal"
                subvia = 602
            Case "óvulo"
                subvia = 602
            Case "pomada vaginal"
                subvia = 602
            Case "solução vaginal"
                subvia = 602
            Case "creme vaginal + óvulo"
                subvia = 602
            Case "dispositivo intra-uterino"
                subvia = 603

                'infarmed descodificou do 603
            Case "dispositivo de libertação intra-uterino"
                subvia = 604
            Case "comprimido para suspensão rectal"
                subvia = 700
            Case "enema, solução"
                subvia = 700
            Case "espuma rectal"
                subvia = 700
            Case "pomada rectal"
                subvia = 700
            Case "pomada rectal + supositório"
                subvia = 700
            Case "solução rectal"
                subvia = 700
            Case "suspensão rectal"
                subvia = 700
            Case "enema, suspensão"
                subvia = 700
            Case "supositório"
                subvia = 704
            Case "cápsula dura + pó e solvente para solução injectável"
                subvia = 800
            Case "concentrado e solvente para solução para perfusão"
                subvia = 800
            Case "concentrado para solução injectável"
                subvia = 800
            Case "concentrado para solução injectável ou para perfusão"
                subvia = 800
            Case "concentrado para solução para perfusão"
                subvia = 800
            Case "emulsão injectável"
                subvia = 800
            Case "emulsão para perfusão"
                subvia = 800
            Case "liofilizado para solução para perfusão"
                subvia = 800
            Case "pó e solvente para solução injectável"
                subvia = 800
            Case "pó e solvente para solução injectável ou para perfusão"
                subvia = 800
            Case "pó e solvente para solução para perfusão"
                subvia = 800
            Case "pó e solvente para suspensão injectável"
                subvia = 800
            Case "pó e veículo para suspensão injectável"
                subvia = 800
            Case "pó para concentrado para solução injectável ou para perfusão"
                subvia = 800
            Case "pó para concentrado para solução para perfusão"
                subvia = 800
            Case "pó para solução injectável"
                subvia = 800
            Case "pó para solução injectável ou para perfusão"
                subvia = 800
            Case "pó para solução ou para suspensão injectável"
                subvia = 800
            Case "pó para solução para perfusão"
                subvia = 800
            Case "solução injectável"
                subvia = 800
            Case "solução injectável ou concentrado para solução para perfusão"
                subvia = 800

            Case "solvente/veículo para uso parentérico"
                subvia = 800
            Case "suspensão injectável"
                subvia = 800

            Case "emulsão injectável ou para perfusão"
                subvia = 800

            Case "solução injectável ou para perfusão"
                subvia = 800
            Case "solução para perfusão"
                subvia = 800

            Case "pó para suspensão para implantação"
                subvia = 803
            Case "solução para diálise peritoneal"
                subvia = 806
            Case "comprimido + supositório"
                subvia = 900
            Case "implante"
                subvia = 901
            Case "verniz medicamentoso para as unhas"
                subvia = 902
            Case "verniz para as unhas medicamentoso"
                subvia = 902
272:        Case "seringa"
273:
                subvia = 21
274:        Case "agulhas"
275:
                subvia = 22
276:        Case "lancetas"

277:            subvia = 23
278:        Case "lancetas para punção capilar"

                subvia = 23
279:        Case "lancetas para amostragem de sangue"

                subvia = 23
280:        Case "lancetas esterilizadas por irradiação, de uso único"

281:            subvia = 23
283:        Case "lancetas estéreis"

284:            subvia = 23
285:        Case "lancetas estéreis para obtenção de uma gota de sangue"

286:            subvia = 23
287:        Case "lancetas estéreis por radiação gama"

288:            subvia = 23
289:        Case "lancetas esterilizadas por radiação gama"

290:            subvia = 23
291:        Case "seringa de uso único, estéril para administração de insulina com agulha ultrafina, 30g, 8mm"

292:            subvia = 21
293:        Case "seringas de 0,3 ml (8mm) diametro 0,3mm (30g) escala de 30 unidades divididas em 1/2"

294:            subvia = 21
295:        Case "tiras para determinação de glicémia"

296:            subvia = 24
297:        Case "tiras para determinação de glicosúria"

298:            subvia = 25
299:        Case "tiras para determinação de glicosúria e cetonúria"

300:            subvia = 26
301:        Case "tiras teste de b-cetonemia"

302:            subvia = 27
303:        Case "agulha de 0,30x8mm (30 gx5/16)"

304:            subvia = 22
305:        Case "agulha de 0,33x12mm (29 gx15/32)"

306:            subvia = 22
307:        Case "agulha de uso único, estéril, para canetas de administração de insulina"

308:            subvia = 22



            Case "cápsula dura + pó e solvente para solução injectável"
                subvia = 1010
            Case "cápsula dura de libertação modificada"
                subvia = 102
            Case "cápsula dura de libertação prolongada"
                subvia = 102
            Case "cápsula dura gastrorresistente"
                subvia = 101
            Case "cápsula dura"
                subvia = 101
            Case "cápsula mole de libertação modificada"
                subvia = 102
            Case "cápsula mole gastrorresistente"
                subvia = 101
            Case "cápsula para inalação por vaporização"
                subvia = 5
            Case "cápsula para inalação"
                subvia = 5
            Case "colírio de acção prolongada"
                subvia = 9
            Case "colírio, pó e solvente para solução"
                subvia = 9
            Case "comprimido + supositório"
                subvia = 0
            Case "comprimido revestido + comprimido"
                subvia = 101
            Case "concentrado para solução injectável ou para perfusão"
                subvia = 12
            Case "dispositivo intra-uterino"
                subvia = 7
            Case "emulsão injectável ou para perfusão"
                subvia = 12
            Case "emulsão injectável"
                subvia = 12
            Case "emulsão para perfusão"
                subvia = 12
            Case "enema, solução"
                subvia = 13
            Case "enema, suspensão"
                subvia = 13
            Case "gás para inalação"
                subvia = 5
            Case "gel dental"
                subvia = 2
            Case "goma para mascar medicamentosa"
                subvia = 1011
            Case "granulado para xarope"
                subvia = 107
            Case "liofilizado para solução para perfusão"
                subvia = 12
            Case "líquido para inalação por vaporização"
                subvia = 5
            Case "pasta dentífrica"
                subvia = 2
            Case "pó e solvente para solução injectável ou para perfusão"
                subvia = 12
            Case "pó e solvente para solução oral"
                subvia = 107
            Case "pó e solvente para suspensão injectável"
                subvia = 12
            Case "pó e solvente para suspensão oral"
                subvia = 107
            Case "pó para concentrado para solução injectável ou para perfusão"
                subvia = 12
            Case "pó para concentrado para solução para perfusão"
                subvia = 12
            Case "pó para inalação, cápsula dura"
                subvia = 5
            Case "pó para suspensão para implantação"
                subvia = 0
            Case "solução injectável ou concentrado para solução para perfusão"
                subvia = 12
            Case "solução para diálise peritoneal"
                subvia = 0
            Case "solução para pulverização auricular"
                subvia = 0
            Case "suspensão dental"
                subvia = 2
            Case "suspensão"
                subvia = 107
            Case "verniz medicamentoso para as unhas"
                subvia = 3
            Case "adesivo cutâneo"
                subvia = 999
            Case "colírio, solução em recipiente unidose"
                subvia = 999
            Case "comprimido gastrorresistente de libertação prolongada"
                subvia = 999
            Case "concentrado para solução oral"
                subvia = 999
            Case "emplastro para teste cutâneo"
                subvia = 999
            Case "gel bucal"
                subvia = 999
            Case "gel intestinal"
                subvia = 999
            Case "gel periodontal"
                subvia = 999
            Case "gotas auriculares ou colírio, solução"
                subvia = 999
            Case "granulado gastrorresistente de libertação prolongada"
                subvia = 999
            Case "granulado revestido em saqueta"
                subvia = 999
            Case "pastilha mole"
                subvia = 999
            Case "película bucal"
                subvia = 999
            Case "penso(inpregnado, saqueta)"
                subvia = 999
            Case "pó e solução para solução injectável"
                subvia = 999
            Case "pó e solvente para solução cutânea"
                subvia = 999
            Case "pó e solvente para solução para inalação por nebulização"
                subvia = 999
            Case "pó e suspensão para suspensão injectável"
                subvia = 999
            Case "pó para solução oral em saqueta"
                subvia = 999
            Case "pó para suspensão injectável + suspensão injectável"
                subvia = 999
            Case "pó para suspensão oral ou rectal"
                subvia = 999
            Case "pomada nasal"
                subvia = 999
            Case "solução dental"
                subvia = 999
            Case "solução injectável em seringa pré-cheia"
                subvia = 999
            Case "solução para irrigação"
                subvia = 999
            Case "solução para pulverização bucal ou nasal"
                subvia = 999
            Case "suspensão cutânea"
                subvia = 999
            Case "suspensão injectável em seringa pré-cheia"
                subvia = 999




309:        Case Else
310:            MsgBox("forma farmacêutica (subvia) desconhecida")
511:            End Select
512:    If subvia = 0 Then
513:        MsgBox("subvia = 0")
514:    End If
515:    Exit Function
mostrarerro:
        MsgBox("function subvia: erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    'Private Sub butForm3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butForm3.Click
    '       On Error GoTo MOSTRARERRO
    '1:      Form3.BringToFront()
    '        Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    '   End Sub


    'retirado quando introduzi o menu principal
    'Private Sub butEC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEC.Click 
    '   On Error GoTo MOSTRARERRO
    '1:      EC.Show()
    '       Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub


    'retirado quando introduzi o menu principal
    'Private Sub butEP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEP.Click
    '   On Error GoTo MOSTRARERRO
    '   EP.Show()
    '   Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub


    Sub portariado(ByVal qual As Short, ByVal port As Short)
1:      On Error GoTo MOSTRARERRO
2:      haport = True
3:      Select Case qual
            Case Is = 1
4:              aviam1.BackColor = Color.Beige
5:              Select Case port
                    Case Is = "1474"
7:                      but1474_01.Visible = True
8:                  Case Is = "1234"
9:                      but1234_01.Visible = True
                        'MsgBox(a1row(1) & " necessita de especialidade" & vbCr _
                        '   & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                        '   & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
13:                     aviam1.BackColor = Color.Purple
14:                 Case Is = "14123"
15:                     but14123_01.Visible = True
16:                 Case Is = "4250"
17:                     but4250_01.Visible = True
18:                     'MsgBox(a1row(1) & " necessita de" & vbCr _
                        '' & "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                        ''   & "e especialidade de neurologia ou psiquiatria" & vbCr _
                        '   & "para ser comparticipado")
19:                     aviam1.BackColor = Color.Purple
20:                 Case Is = "21094"
21:                     but21094_01.Visible = True
22:                     '               MsgBox(a1row(1) & " necessita de especialidade" & vbCr _
                        '  & "de psiquiatria ou de neurologia" & vbCr _
                        '& "no caso de comparticipado com o despacho nº. 21094/99")
23:                     aviam1.BackColor = Color.Purple
24:                 Case Is = "10910"
25:                     but10910_01.Visible = True
26:                     '            MsgBox(a1row(0) & " necessita de especialidade" & vbCr _
                        '& "de reumatologia ou de medicina interna" & vbCr _
                        '& "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
27:                     aviam1.BackColor = Color.Purple
28:                 Case Is = "10279"
29:                     but10279_01.Visible = True
30:                 Case Is = "62010"
31:                     but62010_01.Visible = True
32:                     End Select
33:         Case Is = 2
34:             aviam2.BackColor = Color.Beige
35:             Select Case port
                    Case Is = "1474"
37:                     but1474_02.Visible = True
38:                 Case Is = "1234"
39:                     but1234_02.Visible = True
40:                     '              MsgBox(a2row(1) & " necessita de especialidade" & vbCr _
                        '    & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                        '   & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
41:                     aviam2.BackColor = Color.Purple
42:                 Case Is = "14123"
43:                     but14123_02.Visible = True
44:                 Case Is = "4250"
45:                     but4250_02.Visible = True
46:                     '              MsgBox(a2row(1) & " necessita de" & vbCr _
                        '     & "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                        '  & "e especialidade de neurologia ou psiquiatria" & vbCr _
                        '  & "para ser comparticipado")
47:                     aviam2.BackColor = Color.Purple
48:                 Case Is = "21094"
49:                     but21094_02.Visible = True
50:                     '           MsgBox(a2row(1) & " necessita de especialidade" & vbCr _
                        '  & "de psiquiatria ou de neurologia" & vbCr _
                        '  & "no caso de comparticipado com o despacho nº. 21094/99")
51:                     aviam2.BackColor = Color.Purple
52:                 Case Is = "10910"
53:                     but10910_02.Visible = True
54:                     '            MsgBox(a2row(0) & " necessita de especialidade" & vbCr _
                        ' & "de reumatologia ou de medicina interna" & vbCr _
                        ' & "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
55:                     aviam2.BackColor = Color.Purple
56:                 Case Is = "10279"
57:                     but10279_02.Visible = True
58:                 Case Is = "62020"
59:                     but62010_02.Visible = True
60:                     End Select
61:         Case Is = 3
62:             aviam3.BackColor = Color.Beige
63:             Select Case port
                    Case Is = "1474"
65:                     but1474_03.Visible = True
66:                 Case Is = "1234"
67:                     but1234_03.Visible = True
68:                     '            MsgBox(a3row(1) & " necessita de especialidade" & vbCr _
                        '   & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                        '   & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
69:                     aviam3.BackColor = Color.Purple
70:                 Case Is = "14123"
71:                     but14123_03.Visible = True
72:                 Case Is = "4250"
73:                     but4250_03.Visible = True
74:                     '              MsgBox(a3row(1) & " necessita de" & vbCr _
                        '    & "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                        '    & "e especialidade de neurologia ou psiquiatria" & vbCr _
                        '   & "para ser comparticipado")
75:                     aviam3.BackColor = Color.Purple
79:                 Case Is = "21094"
77:                     but21094_03.Visible = True
78:                     '             MsgBox(a3row(1) & " necessita de especialidade" & vbCr _
                        '    & "de psiquiatria ou de neurologia" & vbCr _
                        '   & "no caso de comparticipado com o despacho nº. 21094/99")
80:                     aviam3.BackColor = Color.Purple
81:                 Case Is = "10910"
82:                     but10910_03.Visible = True
83:                     '           MsgBox(a3row(0) & " necessita de especialidade" & vbCr _
                        '& "de reumatologia ou de medicina interna" & vbCr _
                        '& "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
84:                     aviam3.BackColor = Color.Purple
85:                 Case Is = "10279"
86:                     but10279_03.Visible = True
87:                 Case Is = "62030"
88:                     but62010_03.Visible = True
89:                     End Select
90:         Case Is = 4
91:             aviam4.BackColor = Color.Beige
92:             Select Case port
                    Case Is = "1474"
94:                     but1474_04.Visible = True
95:                 Case Is = "1234"
96:                     but1234_04.Visible = True
97:                     '          MsgBox(a4row(1) & " necessita de especialidade" & vbCr _
                        ' & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                        ' & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
98:                     aviam4.BackColor = Color.Purple
99:                 Case Is = "14123"
100:                    but14123_04.Visible = True
101:                Case Is = "4250"
102:                    but4250_04.Visible = True
103:                    '          MsgBox(a4row(1) & " necessita de" & vbCr _
                        '& "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                        '& "e especialidade de neurologia ou psiquiatria" & vbCr _
                        '& "para ser comparticipado")
104:                    aviam4.BackColor = Color.Purple
105:                Case Is = "21094"
106:                    but21094_04.Visible = True
107:                    '         MsgBox(a4row(1) & " necessita de especialidade" & vbCr _
                        '& "de psiquiatria ou de neurologia" & vbCr _
                        '& "no caso de comparticipado com o despacho nº. 21094/99")
108:                    aviam4.BackColor = Color.Purple
109:                Case Is = "10910"
110:                    but10910_04.Visible = True
111:                    '           MsgBox(a4row(0) & " necessita de especialidade" & vbCr _
                        '& "de reumatologia ou de medicina interna" & vbCr _
                        '& "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
112:                    aviam4.BackColor = Color.Purple
113:                Case Is = "10279"
114:                    but10279_04.Visible = True
115:                Case Is = "62040"
116:                    but62010_04.Visible = True
117:                    End Select
118:            'Case Is = 5
119:            '   aviam5.BackColor = Color.Beige
120:            '   Select Case port
                ' Case Is = "1474"
122:            '     but1474_05.Visible = True
123:            ' Case Is = "1234"
124:            '     but1234_05.Visible = True
125:            '     MsgBox(a5row(1) & " necessita de especialidade" & vbCr _
                '& "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                '& "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
126:            '         aviam5.BackColor = Color.Purple
127:            '     Case Is = "14123"
128:            '         but14123_05.Visible = True
129:            '     Case Is = "4250"
130:            '         but4250_05.Visible = True
131:            '         MsgBox(a5row(1) & " necessita de" & vbCr _
                '& "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                '& "e especialidade de neurologia ou psiquiatria" & vbCr _
                '& "para ser comparticipado")
132:            '         aviam5.BackColor = Color.Purple
133:            '    Case Is = "21094"
134:            '       but21094_05.Visible = True
135:            '      MsgBox(a5row(1) & " necessita de especialidade" & vbCr _
                ' & "de psiquiatria ou de neurologia" & vbCr _
                ' & "no caso de comparticipado com o despacho nº. 21094/99")
136:            '            aviam5.BackColor = Color.Purple
137:            '       Case Is = "10910"
138:            '          but10910_05.Visible = True
139:            '         MsgBox(a5row(0) & " necessita de especialidade" & vbCr _
                '& "de reumatologia ou de medicina interna" & vbCr _
                '& "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
140:            '           aviam5.BackColor = Color.Purple
141:            '      Case Is = "10279"
142:            '         but10279_05.Visible = True
143:            '    Case Is = "62050"
144:            '       but62010_05.Visible = True
145:            '      End Select
146:            'Case Is = 6
147:            '   aviam6.BackColor = Color.Beige
148:            '  Select Case port
                'Case Is = "1474"
150:            '  but1474_06.Visible = True
151:            'Case Is = "1234"
152:            '   but1234_06.Visible = True
153:            ' MsgBox(a6row(1) & " necessita de especialidade" & vbCr _
                '& "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
                '& "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
154:            '        aviam6.BackColor = Color.Purple
155:            '   Case Is = "14123"
156:            '      but14123_06.Visible = True
157:            ' Case Is = "4250"
158:            '    but4250_06.Visible = True
159:            '   MsgBox(a6row(1) & " necessita de" & vbCr _
                '& "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
                '& "e especialidade de neurologia ou psiquiatria" & vbCr _
                '& "para ser comparticipado")
160:            '          aviam6.BackColor = Color.Purple
161:            '      Case Is = "21094"
162:            '          but21094_06.Visible = True
163:            '          MsgBox(a6row(1) & " necessita de especialidade" & vbCr _
                ' & "de psiquiatria ou de neurologia" & vbCr _
                ' & "no caso de comparticipado com o despacho nº. 21094/99")
164:            '         aviam6.BackColor = Color.Purple
165:            '     Case Is = "10910"
166:            '          but10910_06.Visible = True
167:            '           MsgBox(a6row(0) & " necessita de especialidade" & vbCr _
                '& "de reumatologia ou de medicina interna" & vbCr _
                '& "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
168:            '          aviam6.BackColor = Color.Purple
169:            '     Case Is = "10279"
170:            '        but10279_06.Visible = True
171:            '   Case Is = "62010"
172:            '      but62010_06.Visible = True
173:            '     End Select
174:            End Select
175:    Exit Sub
MOSTRARERRO:
        MsgBox("sub portariado: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Private Sub codEC2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles codEC2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta44 As String
2:      Caracta44 = codEC2.Text
3:      If Len(codEC2.Text) = 7 Then
4:          If Caracta44 Like "#######" Then
5:              limpar4()
6:              codigo4.codigo = codEC2.Text
7:              incorporar()
8:              mostrar()
9:              Me.CodeEC1.Focus()
10:             CodeEC1.SelectionStart = 0
11:             CodeEC1.SelectionLength = Len(CodeEC1.Text)
12:         End If
13:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub codEC2_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub codeEC1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CodeEC1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta441 As String
2:      Caracta441 = CodeEC1.Text
3:      If Len(CodeEC1.Text) = 8 Then
4:          If Caracta441 Like "########" Then
5:              limpar41()
6:              codigo41.codigo = CodeEC1.Text
7:              incorporar1()
8:              mostrar1()
9:              Me.codEC2.Focus()
10:             codEC2.SelectionStart = 0
11:             codEC2.SelectionLength = Len(codEC2.Text)
12:         End If
13:     End If
14:     If Len(CodeEC1.Text) = 7 Then
16:         If Caracta441 Like "#######" Then
18:             limpar41()
20:             codigo41.codigo = CodeEC1.Text
22:             incorporar1()
24:             mostrar1()
26:             Me.codEC2.Focus()
28:             codEC2.SelectionStart = 0
30:             codEC2.SelectionLength = Len(codEC2.Text)
32:         End If
33:     End If
34:     Exit Sub
MOSTRARERRO:
        MsgBox("Sub codeEC1_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub limpar4()
        On Error GoTo MOSTRARERRO
1:      mostrado9 = "false"
2:      portec4a = ""
3:      portec4b = ""
4:      genec4 = ""
5:      mostranome = ""
6:      mostradci = ""
7:      mostraforma = ""
8:      mostradose = ""
        ghtext = ""
        mostraqty = ""
        mostragh = ""
9:      mostracomp = ""
10:     mostracompgen = ""
        labelmedcnpem.Text = ""
        labelmedcnpem.BackColor = Color.Transparent
        labelmedcnpem.ForeColor = Color.Black
        labelmedcnpem.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Bold)
11:     labelmednome.Text = ""
12:     labelmeddci.Text = ""
13:     labelmedforma.Text = ""
14:     labelmeddose.Text = ""
        labelmedqty.Text = ""
        labelmedgh.Text = ""
15:     labelmedcompgen.Text = ""
16:     labelmednome.BackColor = Color.Transparent
17:     labelmednome.ForeColor = Color.Black
18:     labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Regular)
19:     labelmedcompgen.Text = ""
20:     labelmedcompgen.BackColor = Color.Transparent
21:     labelmedcompgen.ForeColor = Color.Black
22:     labelmedcompgen.Font = New Font(Me.labelmedcompgen.Font, FontStyle.Regular)
23:     labelmeddose.Text = ""
24:     labelmeddose.BackColor = Color.Transparent
25:     labelmeddose.ForeColor = Color.Black
26:     labelmeddose.Font = New Font(Me.labelmeddose.Font, FontStyle.Regular)
27:     labelmedqty.Text = ""
        labelmedqty.BackColor = Color.Transparent
        labelmedqty.ForeColor = Color.Black
        labelmedqty.Font = New Font(Me.labelmedqty.Font, FontStyle.Regular)
        labelmedgh.Text = ""
        labelmedgh.BackColor = Color.Transparent
        labelmedgh.ForeColor = Color.Black
        labelmedgh.Font = New Font(Me.labelmedgh.Font, FontStyle.Regular)
        labelmeddci.Text = ""
28:     labelmeddci.BackColor = Color.Transparent
29:     labelmeddci.ForeColor = Color.Black
30:     labelmeddci.Font = New Font(Me.labelmeddci.Font, FontStyle.Regular)
31:     labelmedforma.Text = ""
32:     labelmedforma.BackColor = Color.Transparent
33:     labelmedforma.ForeColor = Color.Black
34:     labelmedforma.Font = New Font(Me.labelmedforma.Font, FontStyle.Regular)
35:     labelmedports.Text = ""
36:     labelmedports.BackColor = Color.Transparent
37:     labelmedports.ForeColor = Color.Black
38:     labelmedports.Font = New Font(Me.labelmedports.Font, FontStyle.Regular)
39:     labelmedports.Text = ""
40:     mostraports = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limpar4: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub limpar41()
        On Error GoTo MOSTRARERRO
1:      mostrado91 = "false"
2:      portec4a1 = ""
3:      portec4b1 = ""
4:      genec41 = ""
        mostracnpem1 = ""
5:      mostranome1 = ""
6:      mostradci1 = ""
7:      mostraforma1 = ""
8:      mostradose1 = ""
        ghtext1 = ""
        mostraqty1 = ""
        mostragh1 = ""
9:      mostracomp1 = ""
10:     mostracompgen1 = ""
40:     mostraports1 = ""
        labelmedcnpem1.Text = ""
        labelmednome1.Text = ""
        labelmeddci1.Text = ""
        labelmedforma1.Text = ""
        labelmeddose1.Text = ""
        labelmedqty1.Text = ""
        labelmedgh1.Text = ""
        labelmedcompgen1.Text = ""
        labelmednome1.BackColor = Color.Transparent
        labelmednome1.ForeColor = Color.Black
        labelmednome1.Font = New Font(Me.labelmednome.Font, FontStyle.Regular)
        labelmedcnpem1.BackColor = Color.Transparent
        labelmedcnpem1.ForeColor = Color.Black
        labelmedcnpem1.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Regular)
        labelmedcompgen1.Text = ""
        labelmedcompgen1.BackColor = Color.Transparent
        labelmedcompgen1.ForeColor = Color.Black
        labelmedcompgen1.Font = New Font(Me.labelmedcompgen.Font, FontStyle.Regular)
        labelmeddose1.Text = ""
        labelmeddose1.BackColor = Color.Transparent
        labelmeddose1.ForeColor = Color.Black
        labelmeddose1.Font = New Font(Me.labelmeddose.Font, FontStyle.Regular)
        labelmedqty1.Text = ""
        labelmedqty1.BackColor = Color.Transparent
        labelmedqty1.ForeColor = Color.Black
        labelmedqty1.Font = New Font(Me.labelmedqty.Font, FontStyle.Regular)
        labelmedgh1.Text = ""
        labelmedgh1.BackColor = Color.Transparent
        labelmedgh1.ForeColor = Color.Black
        labelmedgh1.Font = New Font(Me.labelmedgh.Font, FontStyle.Regular)
        labelmeddci1.Text = ""
        labelmeddci1.BackColor = Color.Transparent
        labelmeddci1.ForeColor = Color.Black
        labelmeddci1.Font = New Font(Me.labelmeddci.Font, FontStyle.Regular)
        labelmedforma1.Text = ""
        labelmedforma1.BackColor = Color.Transparent
        labelmedforma1.ForeColor = Color.Black
        labelmedforma1.Font = New Font(Me.labelmedforma.Font, FontStyle.Regular)
        labelmedports1.Text = ""
        labelmedports1.BackColor = Color.Transparent
        labelmedports1.ForeColor = Color.Black
        labelmedports1.Font = New Font(Me.labelmedports.Font, FontStyle.Regular)
        labelmedports1.Text = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limpar41: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub incorporar()
        On Error GoTo MOSTRARERRO
1:      If novado = "false" Then
2:          irbuscar4()
3:      End If
4:
5:      If Not IsNothing(codigorow) Then
6:          If codigorow(8) = True Then
7:              genec4 = "genérico"
8:          Else
9:              genec4 = "de marca"
10:         End If
11:
12:         If codigorow(10) = True Then
13:             portec4a = "despacho 4250/2007"
14:             portec4b = ""
15:         Else
16:             If codigorow(13) = True Then
17:                 portec4a = "despacho 10279/2008"
18:                 portec4b = "despacho 10280/2008"
19:             Else
20:                 If codigorow(17) = True Then
                        If port1474 = True Then
21:                         '     portec4a = "portaria 1474/2004 (ad)"
22:                         '    portec4b = ""
                        End If
23:                 Else
24:                     If codigorow(18) = True Then
                            If port1474 = True Then
25:                             '     portec4a = "portaria 1474/2004 (nl)"
26:                             '    portec4b = ""
                            End If
27:                     Else
28:                         If codigorow(12) = True Then
29:                             portec4a = "despacho 21094/1999"
30:                             portec4b = ""
31:                         Else
32:                             If codigorow(15) = True Then
33:                                 portec4a = "despacho 10910/2009"
34:                                 portec4b = ""
35:                             Else
a36:                                If codigorow(19) = True Then
a37:                                    portec4a = "lei 6/2010"
a38:                                    portec4b = ""
a39:                                Else
36:                                     If codigorow(11) = True Then
37:                                         portec4a = "despacho 1234/2007"
38:                                         If codigorow(16) = True Then
39:                                             portec4b = "despacho 14123/2009"
40:                                         Else
41:                                             portec4b = ""
42:                                         End If
43:                                     End If
44:                                 End If
45:                             End If
46:                         End If
47:                     End If
48:                 End If
49:             End If
50:         End If
            labelmedcnpem.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Bold)
51:         labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
52:         labelmeddci.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
53:         labelmedforma.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
54:         labelmeddose.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedqty.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedgh.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
55:         labelmedcompgen.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
56:         labelmedports.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
57:         mostranome = (codigorow(2))
58:         mostradci = (codigorow(1))
59:         mostraforma = (codigorow(3)) & " " & via(codigorow(3))
60:         mostradose = (codigorow(4))
            mostracnpem = (codigorow(22))
61:         If IsNumeric(codigorow(5)) Then
62:             mostraqty = (codigorow(5)) & " unidade(s)"
63:         Else : mostraqty = (codigorow(5))
64:
65:         End If
66:         Select Case (codigorow(7))
                Case Is = 0
68:                 ghtext = ""
69:             Case 1 To 9
70:                 ghtext = "GH000"
71:             Case 10 To 99
                    ghtext = "GH00"
72:             Case 100 To 999
73:                 ghtext = "GH0"
74:             Case Else
75:                 ghtext = "GH"
76:                 End Select
77:         mostragh = codigorow(23) & " / " & codigorow(25) & " / " & codigorow(21) & " ; " & ghtext & (codigorow(7))
78:
79:         mostracompgen = (codigorow(6)) & "% (" & (genec4) & " = " & (codigorow(9)) & ")"
80:         mostraports = "(" & (portec4a) & ") (" & (portec4b) & ")"
81:         If novado = "false" Then
82:         End If
83:     ElseIf mostrado9 = False Then
84:         aviadoNexist(9)
85:         mostrado9 = "True"
86:     End If
87:     Exit Sub
MOSTRARERRO:
        MsgBox("Sub incorporar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub incorporar1()
        ' On Error GoTo MOSTRARERRO
1:      If novado1 = "false" Then
2:          irbuscar41()
3:      End If
4:
5:      If Not IsNothing(codigorow) Then
6:          If codigorow(8) = True Then
7:              genec41 = "genérico"
8:          Else
9:              genec41 = "de marca"
10:         End If
11:
12:         If codigorow(10) = True Then
13:             portec4a1 = "despacho 4250/2007"
14:             portec4b1 = ""
15:         Else
16:             If codigorow(13) = True Then
17:                 portec4a1 = "despacho 10279/2008"
18:                 portec4b1 = "despacho 10280/2008"
19:             Else
20:                 If codigorow(17) = True Then
                        If port1474 = True Then
21:                         '    portec4a1 = "portaria 1474/2004 (ad)"
22:                         '    portec4b1 = ""
                        End If
23:                 Else
24:                     If codigorow(18) = True Then
                            If port1474 = True Then
25:                             '       portec4a1 = "portaria 1474/2004 (nl)"
26:                             '      portec4b1 = ""
                            End If
27:                     Else
28:                         If codigorow(12) = True Then
29:                             portec4a1 = "despacho 21094/1999"
30:                             portec4b1 = ""
31:                         Else
32:                             If codigorow(15) = True Then
33:                                 portec4a1 = "despacho 10910/2009"
34:                                 portec4b1 = ""
35:                             Else
a36:                                If codigorow(19) = True Then
a37:                                    portec4a1 = "lei 6/2010"
a38:                                    portec4b1 = ""
a39:                                Else
36:                                     If codigorow(11) = True Then
37:                                         portec4a1 = "despacho 1234/2007"
38:                                         If codigorow(16) = True Then
39:                                             portec4b1 = "despacho 14123/2009"
40:                                         Else
41:                                             portec4b1 = ""
42:                                         End If
43:                                     End If
44:                                 End If
45:                             End If
46:                         End If
47:                     End If
48:                 End If
49:             End If
50:         End If
            labelmedcnpem1.Width = 100
51:         labelmednome1.Width = 13 * Len(codigorow(2).ToString)
52:         labelmeddci1.Width = 13 * Len(codigorow(1).ToString)
53:         labelmedforma1.Width = 13 * (Len(codigorow(3).ToString) + 4)
54:         labelmeddose1.Width = 13 * Len(codigorow(4).ToString)
55:         If IsNumeric(codigorow(5)) Then
56:             labelmedqty1.Width = 13 * Len(String.Concat(codigorow(5).ToString, " unidade(s)"))
57:         Else : labelmedqty1.Width = 13 * Len(codigorow(5).ToString)
58:
59:         End If
60:         labelmedcnpem1.Width = 100
61:         labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (codigorow(9).ToString) & ") " & codigorow(6).ToString & "% "))
62:         labelmedgh1.Width = 250
63:         labelmednome1.Location = New Point(650 - labelmednome1.Width, 492)
64:         labelmeddci1.Location = New Point(650 - labelmeddci1.Width, 522)
65:         labelmedforma1.Location = New Point(650 - labelmedforma1.Width, 552)
66:         labelmeddose1.Location = New Point(650 - labelmeddose1.Width, 582)
67:         labelmedqty1.Location = New Point(650 - labelmedqty1.Width, 612)
68:         labelmedcompgen1.Location = New Point(650 - labelmedcompgen1.Width, 642)
69:         labelmedgh1.Location = New Point(650 - labelmedgh1.Width, 702)
            labelmedcnpem1.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Bold)
70:         labelmednome1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
71:         labelmeddci1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
72:         labelmedforma1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
73:         labelmeddose1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
74:         labelmedqty1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
75:         labelmedgh1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
76:         labelmedcompgen1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
77:         labelmedports1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            mostracnpem1 = (codigorow(22))
78:         mostranome1 = (codigorow(2))
79:         mostradci1 = (codigorow(1))
80:         mostraforma1 = via(codigorow(3)) & " " & codigorow(3)
81:         mostradose1 = (codigorow(4))
82:
83:         If IsNumeric(codigorow(5)) Then
84:             mostraqty1 = (codigorow(5)) & " unidade(s)"
85:         Else : mostraqty1 = (codigorow(5))
86:
87:         End If
88:
89:         Select Case (codigorow(7))
                Case Is = 0
91:                 ghtext1 = ""
92:             Case 1 To 9
93:                 ghtext1 = "GH000"
94:             Case 10 To 99
95:                 ghtext1 = "GH00"
96:             Case 100 To 999
97:                 ghtext1 = "GH0"
98:             Case Else
99:                 ghtext1 = "GH"
100:                End Select
101:        mostragh1 = ghtext1 & (codigorow(7)) & " ; " & codigorow(21) & " / " & codigorow(25) & " / " & codigorow(23)
102:
103:        mostracompgen1 = ("(" & (genec41) & " = " & (codigorow(9)) & ") " & codigorow(6) & "% ")
104:        mostraports1 = "(" & (portec4a1) & ") (" & (portec4b1) & ")"
105:        If novado1 = "false" Then
106:        End If
107:    ElseIf mostrado91 = False Then
108:        aviadoNexist(7)
109:        mostrado91 = "True"
110:    End If
111:
112:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub incorporar1: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub irbuscar4()
        On Error GoTo MOSTRARERRO
        labelmedcnpem.Font = New System.Drawing.Font("Microsoft Sans Serif", 15)
1:      codigorow = DS.infarmed.FindBycode(codigo4.codigo)
2:      If Not IsNothing(codigorow) Then
3:          codigoarray.Add(codigorow)
            labelmedcnpem.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Bold)
4:          labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
5:          labelmeddci.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
6:          labelmedforma.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
7:          labelmeddose.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedqty.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedgh.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
8:          labelmedcompgen.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
9:          labelmedports.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
10:     ElseIf mostrado9 = False Then
11:         aviadoNexist(9)
12:         mostrado9 = "True"
13:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub irbuscar4: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub irbuscar41()
        On Error GoTo MOSTRARERRO
        labelmedcnpem.Font = New System.Drawing.Font("Microsoft Sans Serif", 15)
1:      codigorow = DS.infarmed.FindBycode(codigo41.codigo)
2:      If Not IsNothing(codigorow) Then
3:          codigoarray.Add(codigorow)
            labelmedcnpem1.Font = New Font(Me.labelmedcnpem.Font, FontStyle.Bold)
4:          labelmednome1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
5:          labelmeddci1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
6:          labelmedforma1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
7:          labelmeddose1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedqty1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
            labelmedgh1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
8:          labelmedcompgen1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
9:          labelmedports1.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
10:     ElseIf mostrado91 = False Then
11:         aviadoNexist(7)
12:         mostrado91 = "True"
13:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub irbuscar41: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub mostrar()
        On Error GoTo MOSTRARERRO
1:      If mostrado9 = False Then
            labelmedcnpem.Text = mostracnpem
2:          labelmednome.Text = mostranome
3:          labelmeddci.Text = mostradci
4:          labelmedforma.Text = mostraforma
5:          labelmeddose.Text = mostradose
            labelmedqty.Text = mostraqty
            labelmedgh.Text = mostragh
6:          labelmedcompgen.Text = mostracompgen
7:          labelmedports.Text = mostraports
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub mostrar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Sub mostrar1()
        On Error GoTo MOSTRARERRO
1:      If mostrado91 = False Then
            labelmedcnpem1.Text = mostracnpem1
2:          labelmednome1.Text = mostranome1
3:          labelmeddci1.Text = mostradci1
4:          labelmedforma1.Text = mostraforma1
5:          labelmeddose1.Text = mostradose1
            labelmedqty1.Text = mostraqty1
            labelmedgh1.Text = mostragh1
6:          labelmedcompgen1.Text = mostracompgen1
7:          labelmedports1.Text = mostraports1
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub mostrar1: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    '    Private Sub limparEC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        On Error GoTo MOSTRARERRO
    '1:      limpar4()
    '2:      codEC2.Text = ""
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub


    Sub indicar(ByVal which As Short)

1:      On Error GoTo MOSTRARERRO
        If naoindicar = False Then
2:          If Not IsNothing(codigorow) Then



3:              Select Case which
                    Case 1
                        If valorcombopvp1 = True Then

4:                          If av1.mostrado = "true" Then
5:                              If a1row(8) = "true" Then
6:                                  '        gen1.Text = "genérico"
7:                                  gen = True
8:                              Else
9:                                  '        gen1.Text = "marca"
10:                                 gen = False
11:                             End If
12:                             portaria() 'só devolve texto com qual a portaria que leva
13:                             port1.Text = portimedio 'só devolve texto com qual a portaria que leva
14:                             comp = (a1row(6) * 0.01)
15:                             portcomp1 = portcomp

16:                             intermedio = Convert.ToDouble(ComboBox1.SelectedValue.ToString)
17:                             pvp1.Text = intermedio
18:                             pr1 = Replace(a1row(24), ".", ",")
19:                             If organismo = 48 Or organismo = 49 Then    'útil quando existe PRE
20:                                 pr1 = taxapr * pr1                      ' possível fazer PRE=PR+20%, PRE=PR+25%, etc
21:                             End If
22:                             pr = pr1
23:                             If pr > 0 Then
                                    intermedio = pr    'existindo PR, o PVP não interessa pois o cálculo é sobre o PR directamente
24:                                 'o de baixo era quando era usado PVP se PVP<PR
                                    'intermedio = System.Math.Min(intermedio, pr)
25:                             End If
                                'tempcalc = a1row(23)  usei para limitar aqui a compSNS ao valor do PVP e fazia min(calculo,tempcalc)
26:
                                comp1.Text = calculo(organismo, gen, comp, intermedio, Convert.ToDouble(ComboBox1.SelectedValue.ToString), a1row(17)) '17 é pvp e 15 é top5
                                'comp1.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                                'comp1.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
27:                         End If
                        End If
28:                 Case 2
                        If valorcombopvp2 = True Then
29:                         If av2.mostrado = "true" Then
30:                             If a2row(8) = "true" Then
31:                                 '           gen2.Text = "genérico"
32:                                 gen = True
33:                             Else
34:                                 '          gen2.Text = "marca"
35:                                 gen = False
36:                             End If
37:                             portaria()
38:                             port2.Text = portimedio
39:                             comp = (a2row(6) * 0.01)
40:                             portcomp2 = portcomp
41:                             intermedio = Convert.ToDouble(ComboBox2.SelectedValue)
42:                             pvp2.Text = intermedio
43:                             pr2 = Replace(a2row(24), ".", ",")
44:                             If organismo = 48 Or organismo = 49 Then
45:                                 pr2 = taxapr * pr2
46:                             End If
47:                             pr = pr2
48:                             If pr > 0 Then
                                    intermedio = pr
49:                                 'o de baixo era quando era usado PVP se PVP<PR
                                    'intermedio = System.Math.Min(intermedio, pr)
50:                             End If
                                'tempcalc = a2row(23)
                                comp2.Text = calculo(organismo, gen, comp, intermedio, Convert.ToDouble(ComboBox2.SelectedValue.ToString), a2row(17)) '17 é pvp e 15 é top5
51:                             'comp2.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                                'comp2.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
52:                         End If
                        End If
53:                 Case 3
                        If valorcombopvp3 = True Then
54:                         If av3.mostrado = "true" Then
55:                             If a3row(8) = "true" Then
56:                                 '            gen3.Text = "genérico"
57:                                 gen = True
58:                             Else
59:                                 '           gen3.Text = "marca"
60:                                 gen = False
61:                             End If
62:                             portaria()
63:                             port3.Text = portimedio
64:                             comp = (a3row(6) * 0.01)
65:                             portcomp3 = portcomp
66:                             intermedio = Convert.ToDouble(ComboBox3.SelectedValue)
67:                             pvp3.Text = intermedio
68:                             pr3 = Replace(a3row(24), ".", ",")
69:                             If organismo = 48 Or organismo = 49 Then
70:                                 pr3 = taxapr * pr3
71:                             End If
72:                             pr = pr3
73:                             If pr > 0 Then
                                    intermedio = pr
                                    'o de baixo era quando era usado PVP se PVP<PR
74:                                 'intermedio = System.Math.Min(intermedio, pr)
75:                             End If
                                'tempcalc = a3row(23)
                                comp3.Text = calculo(organismo, gen, comp, intermedio, Convert.ToDouble(ComboBox3.SelectedValue.ToString), a3row(17)) '17 é pvp e 15 é top5
76:                             'comp3.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                                'comp3.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
77:                         End If
                        End If
78:                 Case 4
                        If valorcombopvp4 = True Then
79:                         If av4.mostrado = "true" Then
80:                             If a4row(8) = "true" Then
81:                                 '          gen4.Text = "genérico"
82:                                 gen = True
83:                             Else
84:                                 '         gen4.Text = "marca"
85:                                 gen = False
86:                             End If
87:                             portaria()
88:                             port4.Text = portimedio
89:                             comp = (a4row(6) * 0.01)
90:                             portcomp4 = portcomp
91:                             intermedio = Convert.ToDouble(ComboBox4.SelectedValue)
92:                             pvp4.Text = intermedio
93:                             pr4 = Replace(a4row(24), ".", ",")
94:                             If organismo = 48 Or organismo = 49 Then
95:                                 pr4 = taxapr * pr4
96:                             End If
97:                             pr = pr4
98:                             If pr > 0 Then
                                    intermedio = pr
                                    'o de baixo era quando era usado PVP se PVP<PR
99:                                 'intermedio = System.Math.Min(intermedio, pr)
100:                            End If
                                'tempcalc = a4row(23)
                                comp4.Text = calculo(organismo, gen, comp, intermedio, Convert.ToDouble(ComboBox4.SelectedValue.ToString), a4row(17)) '17 é pvp e 15 é top5
101:                            'comp4.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                                'comp4.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
102:                        End If
                        End If
103:                    End Select
154:        End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB indicar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub portaria()
1:      On Error GoTo MOSTRARERRO
2:      If codigorow(10) = True Then
3:          portimedio = "4250"
4:      End If
5:
6:      If codigorow(11) = True Then
7:          portimedio = "1234"
8:      End If
9:
10:     If codigorow(13) = True Then
11:         portimedio = "10279"
12:     End If
13:
14:     If codigorow(14) = True Then
15:         portimedio = "10280"
16:     End If
17:
18:     If codigorow(15) = True Then
19:         portimedio = "10910"
20:     End If
21:
22:     If codigorow(16) = True Then
23:         If codigorow(11) = True Then
24:             portimedio = "1234 + 14123"
25:         Else
26:             portimedio = "14123"
27:         End If
28:     End If
29:
30:     'If codigorow(17) = True Then
        ' If port1474 = True Then
31:     ':     portimedio = "147469"
        '      End If
32:     '  End If
33:
34:     If codigorow(12) = True Then
35:         portimedio = "21094"
36:     End If
37:
38:     'If codigorow(18) = True Then
        ' If port1474 = True Then
39:     ':     portimedio = "1474100"
        '      End If
40:     'End If
41:
42:     If codigorow(19) = True Then
43:         portimedio = "6/2010"
44:     End If

48:     If codigorow(10) = False And codigorow(11) = False And codigorow(13) = False And codigorow(14) = False And codigorow(15) = False _
     And codigorow(16) = False And codigorow(19) = False And codigorow(12) = False Then
            portimedio = "" 'era "não"
49:     End If
50:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB portaria: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Function calculo(ByVal org As Short, ByVal gen As Boolean, ByVal comp As Double, ByVal intermedio As Double, ByVal pvp As Double, ByVal top5 As Double) As Double
1:      On Error GoTo MOSTRARERRO
        org = "1"
2:      If Not IsNothing(org) Then
3:          Select Case org
                Case 1 ', 2 'tipo 10
                    If comp > 0 Then
101:                    If Not IsNothing(intermedio) Then
5:                          calculo = System.Math.Min(System.Math.Round(intermedio * comp, 2), pvp)
102:                    End If
                    Else
                        calculo = 0
                    End If
6:              Case 46 'tipo 17
7:                  If Not IsNothing(intermedio) Then
103:                    calculo = System.Math.Min(System.Math.Round(intermedio * comp, 2), pvp)
104:                End If
8:              Case 42 'tipo 12
9:                  If Not IsNothing(intermedio) Then
105:                    calculo = System.Math.Round(intermedio, 2)
106:                End If
10:             Case 41 'tipo 11
11:                 If comp > 0 Then
12:                     If Not IsNothing(intermedio) Then
107:                        calculo = System.Math.Round(intermedio, 2)
108:                    End If
13:                 Else
14:                     calculo = 0
15:                 End If
16:             Case 67 'tipo 13
17:                 If comp > 0 Then
18:                     If Not IsNothing(intermedio) Then
109:                        calculo = System.Math.Round(intermedio, 2)
110:                    End If
19:                 Else
20:                     calculo = 0
21:                 End If
22:             Case 23, 24, 25 'diabetes - não sei se 24 e 25 também são assim mas já fica
23:                 If Not IsNothing(intermedio) Then
111:                    calculo = System.Math.Round(intermedio * comp, 2)
112:                End If
24:             Case 48 ', 57 'tipo 15
25:                 If comp > 0 Then
                        If pvp <= top5 And top5 > 0 Then
                            If Not IsNothing(intermedio) Then
                                calculo = System.Math.Min(System.Math.Round(tectocomp * intermedio, 2), pvp)
                            End If
                        Else
115:                        If Not IsNothing(intermedio) Then
29:                             calculo = System.Math.Min(System.Math.Round((System.Math.Min(tectocomp, (comp + 0.15))) * intermedio, 2), pvp)
116:                        End If
30:                     End If
31:                 Else
32:                     calculo = 0
33:                 End If
34:             Case 45 ', 59
117:                If Not IsNothing(intermedio) Then
35:                     calculo = System.Math.Min(System.Math.Round(intermedio * (System.Math.Max(comp, portcomp)), 2), pvp)
118:                End If
36:             Case 49 ', 68
37:                 If comp > 0 Then
                        If pvp <= top5 And top5 > 0 Then
                            If Not IsNothing(intermedio) Then
                                calculo = System.Math.Min(System.Math.Round((System.Math.Min(System.Math.Max(tectocomp, (portcomp + 0.15)), tectocomp)) * intermedio, 2), pvp)
                            End If
                        Else
                            If Not IsNothing(intermedio) Then
                                calculo = System.Math.Min(System.Math.Round((System.Math.Min(tectocomp, System.Math.Max((portcomp + 0.15), (comp + 0.15))) * intermedio), 2), pvp)
                            End If
                            : End If
                    Else
                        calculo = 0
                    End If
42:                 'Case 12 'SAMS
123:                '   If Not IsNothing(intermedio) Then
43:                 '                 calculo = intermedio * (0.9)
124:                '                End If
44:                 End Select
45:
46:         'não tenho nada para o tipo 19(47) nem para os organismos
            '13_CGD, 
            '25_SAMSq, 
47:         '09[=(75% - 01)], 
            'CA, 
            'r1, r3, 
            '85, 87, h1[=(90% - 02)], 
            'j1, j7
            'o1, 
            'aa, ab, 
            'xv, 
            'fm, 
            '19
            'ds
            'SF[=(100% - 01)], SG[=(100% - 45)], SH[=(100% - 48)], SI[=(100% - 49)], 
48:     End If
49:     Exit Function
MOSTRARERRO:
        MsgBox("SUB calculo: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function


    Sub verifgen()
1:      On Error GoTo MOSTRARERRO
2:      Dim genverif As Boolean
3:      genverif = False
4:      Select Case A
            Case Is = 1
                If Not IsNothing(a1row) Then
6:                  If a1row(8) = True Then
7:                      genverif = True
8:                  End If
                End If
9:          Case Is = 2
                If Not IsNothing(a2row) And Not IsNothing(a1row) Then
10:                 If a1row(8) = True Or a2row(8) = True Then
11:                     genverif = True
12:                 End If
                End If
13:         Case Is = 3
                If Not IsNothing(a3row) And Not IsNothing(a2row) And Not IsNothing(a1row) Then
14:                 If a1row(8) = True Or a2row(8) = True Or a3row(8) = True Then
15:                     genverif = True
16:                 End If
                End If
17:         Case Is = 4
                If Not IsNothing(a4row) And Not IsNothing(a3row) And Not IsNothing(a2row) And Not IsNothing(a1row) Then
18:                 If a1row(8) = True Or a2row(8) = True Or a3row(8) = True Or a4row(8) = True Then
19:                     genverif = True
20:                 End If
                End If
21:             '     Case Is = 5
22:             '         If a1row(8) = True Or a2row(8) = True Or a3row(8) = True Or a4row(8) = True Or a5row(8) = True Then
23:             ' genverif = True
24:             ' End If
25:             'Case Is = 6
26:             '    If a1row(8) = True Or a2row(8) = True Or a3row(8) = True Or a4row(8) = True Or a5row(8) = True Or a6row(8) = True Then
27:             ' genverif = True
28:             ' End If
29:             End Select
30:     If genverif = True Then
31:         verifgenlabel.Text = "genéricos"
32:         verifgenlabel.BackColor = Color.Yellow
33:     End If
34:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB verifgen: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Function SomarPVP()
1:      On Error GoTo MOSTRARERRO
101:    If naoindicar = False Then
            
2:          Dim somaPVP As Double
3:          If valorcombopvp1 = True Then
                pvp1val = ComboBox1.SelectedValue.ToString
4:              pvp1v = Replace(pvp1val, ".", ",")
5:              pvp1val = Convert.ToDouble(pvp1v)
6:          Else : pvp1val = 0
7:          End If
8:          If valorcombopvp2 = True Then
                pvp2val = ComboBox2.SelectedValue.ToString
9:              pvp2v = Replace(pvp2val, ".", ",")
10:             pvp2val = Convert.ToDouble(pvp2v)
11:         Else : pvp2val = 0
12:         End If
13:         If valorcombopvp3 = True Then
                pvp3val = ComboBox3.SelectedValue.ToString
14:             pvp3v = Replace(pvp3val, ".", ",")
15:             pvp3val = Convert.ToDouble(pvp3v)
16:         Else : pvp3val = 0
17:         End If
18:         If valorcombopvp4 = True Then
                pvp4val = ComboBox4.SelectedValue.ToString
19:             pvp4v = Replace(pvp4val, ".", ",")
20:             pvp4val = Convert.ToDouble(pvp4v)
21:         Else : pvp4val = 0
22:         End If
23:         '   If pvp5.Text <> "" Then
24:         ':     pvp5v = Replace(pvp5.Text, ".", ",")
25:         ':     pvp5val = Convert.ToDouble(pvp5v)
26:         ':     Else : pvp5val = 0
27:         '     End If
28:         '     If pvp6.Text <> "" Then
29:         '     pvp6v = Replace(pvp6.Text, ".", ",")
30:         '    pvp6val = Convert.ToDouble(pvp6v)
31:         '   Else : pvp6val = 0
32:         '  End If
33:         somaPVP = pvp1val + pvp2val + pvp3val + pvp4val ' + pvp5val + pvp6val
34:         SomarPVP = somaPVP
        Else
            SomarPVP = 0
        End If
35:     Exit Function
36:
MOSTRARERRO:
        MsgBox("SUB somarpvp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Function SomarComp()
1:      On Error GoTo MOSTRARERRO
        If naoindicar = False Then
2:          Dim somaComp As Double
3:          If comp1.Text <> "" Then
4:              comp1v = Replace(comp1.Text, ".", ",")
5:              comp1val = Convert.ToDouble(comp1v)
6:          Else : comp1val = 0
7:          End If
8:          If comp2.Text <> "" Then
9:              comp2v = Replace(comp2.Text, ".", ",")
10:             comp2val = Convert.ToDouble(comp2v)
11:         Else : comp2val = 0
12:         End If
13:         If comp3.Text <> "" Then
14:             comp3v = Replace(comp3.Text, ".", ",")
15:             comp3val = Convert.ToDouble(comp3v)
16:         Else : comp3val = 0
17:         End If
18:         If comp4.Text <> "" Then
19:             comp4v = Replace(comp4.Text, ".", ",")
20:             comp4val = Convert.ToDouble(comp4v)
21:         Else : comp4val = 0
22:         End If
23:         '  If comp5.Text <> "" Then
24:         ':     comp5v = Replace(comp5.Text, ".", ",")
25:         ':     comp5val = Convert.ToDouble(comp5v)
26:         ':     Else : comp5val = 0
27:         '     End If
28:         '    If comp6.Text <> "" Then
29:         '     comp6v = Replace(comp6.Text, ".", ",")
30:         '    comp6val = Convert.ToDouble(comp6v)
31:         '   Else : comp6val = 0
32:         '  End If
33:         somaComp = comp1val + comp2val + comp3val + comp4val ' + comp5val + comp6val
34:         SomarComp = somaComp
        Else
            SomarComp = 0
        End If
35:     Exit Function

MOSTRARERRO:
        MsgBox("SUB somarcomp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Sub transferros(ByVal qpa As String)
1:      On Error GoTo MOSTRARERRO
2:
3:      Dim oque As String = "0"
4:      Dim qualpresc As String = "0"
5:      Dim qualav As String = "0"
6:      If qpa > 0 Then
7:          If qpa >= 100 Then
8:              oque = qpa.Substring(0, 1)
9:              qualpresc = qpa.Substring(1, 1)
10:             qualav = qpa.Substring(2, 1)
11:         End If
12:         If qpa < 100 Then
13:             oque = 10
14:             qualpresc = qpa.Substring(0, 1)
15:             qualav = qpa.Substring(1, 1)
16:         End If
17:         If Not IsNothing(qualpresc) Then
18:             Select Case qualpresc
                    Case Is = 1 'Or 5
20:                     qualpresc = p1row(0).ToString
21:                     labelmednome1.Width = 13 * Len(p1row(2).ToString)
22:                     labelmeddci1.Width = 13 * Len(p1row(1).ToString)
23:                     labelmedforma1.Width = 13 * (Len(p1row(3).ToString) + 4)
24:                     labelmeddose1.Width = 13 * Len(p1row(4).ToString)
25:                     If IsNumeric(codigorow(5)) Then
26:                         labelmedqty1.Width = 13 * Len(String.Concat(p1row(5).ToString, " unidade(s)"))
27:                     Else : labelmedqty1.Width = 13 * Len(p1row(5).ToString)
28:
29:                     End If
30:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p1row(9).ToString) & ") " & p1row(6).ToString & "% "))
31:                     labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
32:                 Case Is = 2 'Or 6
33:                     qualpresc = p2row(0).ToString
34:                     labelmednome1.Width = 13 * Len(p2row(2).ToString)
35:                     labelmeddci1.Width = 13 * Len(p2row(1).ToString)
36:                     labelmedforma1.Width = 13 * (Len(p2row(3).ToString) + 4)
37:                     labelmeddose1.Width = 13 * Len(p2row(4).ToString)
38:                     If IsNumeric(codigorow(5)) Then
39:                         labelmedqty1.Width = 13 * Len(String.Concat(p2row(5).ToString, " unidade(s)"))
40:                     Else : labelmedqty1.Width = 13 * Len(p2row(5).ToString)
41:
42:                     End If
43:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p2row(9).ToString) & ") " & p2row(6).ToString & "% "))
44:                     labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
45:                 Case Is = 3 'Or 7
46:                     qualpresc = p3row(0).ToString
47:                     labelmednome1.Width = 13 * Len(p3row(2).ToString)
48:                     labelmeddci1.Width = 13 * Len(p3row(1).ToString)
49:                     labelmedforma1.Width = 13 * (Len(p3row(3).ToString) + 4)
50:                     labelmeddose1.Width = 13 * Len(p3row(4).ToString)
51:                     If IsNumeric(codigorow(5)) Then
52:                         labelmedqty1.Width = 13 * Len(String.Concat(p3row(5).ToString, " unidade(s)"))
53:                     Else : labelmedqty1.Width = 13 * Len(p3row(5).ToString)
54:
55:                     End If
56:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p3row(9).ToString) & ") " & p3row(6).ToString & "% "))
57:                     labelmedgh1.Width = 250
58:                     labelmedcnpem1.Width = 100
59:                 Case Is = 4 'Or 8
60:                     qualpresc = p4row(0).ToString
61:                     labelmednome1.Width = 13 * Len(p4row(2).ToString)
62:                     labelmeddci1.Width = 13 * Len(p4row(1).ToString)
63:                     labelmedforma1.Width = 13 * (Len(p4row(3).ToString) + 4)
64:                     labelmeddose1.Width = 13 * Len(p4row(4).ToString)
65:                     If IsNumeric(codigorow(5)) Then
66:                         labelmedqty1.Width = 13 * Len(String.Concat(p4row(5).ToString, " unidade(s)"))
67:                     Else : labelmedqty1.Width = 13 * Len(p4row(5).ToString)
68:
69:                     End If
70:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p4row(9).ToString) & ") " & p4row(6).ToString & "% "))
71:                     labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
72:                 Case Is = 5
73:                     qualpresc = p1row(0).ToString
74:                     labelmednome1.Width = 13 * Len(p1row(2).ToString)
75:                     labelmeddci1.Width = 13 * Len(p1row(1).ToString)
76:                     labelmedforma1.Width = 13 * (Len(p1row(3).ToString) + 4)
77:                     labelmeddose1.Width = 13 * Len(p1row(4).ToString)
78:                     If IsNumeric(codigorow(5)) Then
79:                         labelmedqty1.Width = 13 * Len(String.Concat(p1row(5).ToString, " unidade(s)"))
80:                     Else : labelmedqty1.Width = 13 * Len(p1row(5).ToString)
81:
82:                     End If
83:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p1row(9).ToString) & ") " & p1row(6).ToString & "% "))
84:                     labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
85:                 Case Is = 6
86:                     qualpresc = p2row(0).ToString
87:                     labelmednome1.Width = 13 * Len(p2row(2).ToString)
88:                     labelmeddci1.Width = 13 * Len(p2row(1).ToString)
89:                     labelmedforma1.Width = 13 * (Len(p2row(3).ToString) + 4)
90:                     labelmeddose1.Width = 13 * Len(p2row(4).ToString)
91:                     If IsNumeric(codigorow(5)) Then
92:                         labelmedqty1.Width = 13 * Len(String.Concat(p2row(5).ToString, " unidade(s)"))
93:                     Else : labelmedqty1.Width = 13 * Len(p2row(5).ToString)
94:
95:                     End If
96:                     labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p2row(9).ToString) & ") " & p2row(6).ToString & "% "))
97:                     labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
98:                 Case Is = 7
99:                     qualpresc = p3row(0).ToString
100:                    labelmednome1.Width = 13 * Len(p3row(2).ToString)
101:                    labelmeddci1.Width = 13 * Len(p3row(1).ToString)
102:                    labelmedforma1.Width = 13 * (Len(p3row(3).ToString) + 4)
103:                    labelmeddose1.Width = 13 * Len(p3row(4).ToString)
104:                    If IsNumeric(codigorow(5)) Then
105:                        labelmedqty1.Width = 13 * Len(String.Concat(p3row(5).ToString, " unidade(s)"))
106:                    Else : labelmedqty1.Width = 13 * Len(p3row(5).ToString)
107:
108:                    End If
109:                    labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p3row(9).ToString) & ") " & p3row(6).ToString & "% "))
110:                    labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
111:                Case Is = 8
112:                    qualpresc = p4row(0).ToString
113:                    labelmednome1.Width = 13 * Len(p4row(2).ToString)
114:                    labelmeddci1.Width = 13 * Len(p4row(1).ToString)
115:                    labelmedforma1.Width = 13 * (Len(p4row(3).ToString) + 4)
116:                    labelmeddose1.Width = 13 * Len(p4row(4).ToString)
117:                    If IsNumeric(codigorow(5)) Then
118:                        labelmedqty1.Width = 13 * Len(String.Concat(p4row(5).ToString, " unidade(s)"))
119:                    Else : labelmedqty1.Width = 13 * Len(p4row(5).ToString)
120:
121:                    End If
122:                    labelmedcompgen1.Width = 13 * Len(String.Concat("(12345678 = " & (p4row(9).ToString) & ") " & p4row(6).ToString & "% "))
123:                    labelmedgh1.Width = 250
                        labelmedcnpem1.Width = 100
124:                    End Select
125:        End If
126:        '260 era aqui
127:        labelmednome1.Location = New Point(650 - labelmednome1.Width, 492)
128:        labelmeddci1.Location = New Point(650 - labelmeddci1.Width, 522)
129:        labelmedforma1.Location = New Point(650 - labelmedforma1.Width, 552)
130:        labelmeddose1.Location = New Point(650 - labelmeddose1.Width, 582)
131:        labelmedqty1.Location = New Point(650 - labelmedqty1.Width, 612)
132:        labelmedcompgen1.Location = New Point(650 - labelmedcompgen1.Width, 642)
133:        labelmedgh1.Location = New Point(650 - labelmedgh1.Width, 702)
134:        If Not IsNothing(qualav) Then
135:            Select Case qualav
                    Case Is = 1
136:                    If Not IsNothing(a1row) Then
137:                        qualav = a1row(0).ToString
138:                    Else
139:                        qualav = 0
140:                    End If
141:                Case Is = 2
142:                    If Not IsNothing(a2row) Then
143:                        qualav = a2row(0).ToString
144:                    Else
145:                        qualav = 0
146:                    End If
147:                Case Is = 3
148:                    If Not IsNothing(a3row) Then
149:                        qualav = a3row(0).ToString
150:                    Else
151:                        qualav = 0
152:                    End If
153:                Case Is = 4
154:                    If Not IsNothing(a4row) Then
155:                        qualav = a4row(0).ToString
156:                    Else
157:                        qualav = 0
158:                    End If
159:                Case Is = 5
160:                    If Not IsNothing(a1row) Then
161:                        qualav = a1row(0).ToString
162:                    Else
163:                        qualav = 0
164:                    End If
165:                Case Is = 6
166:                    If Not IsNothing(a2row) Then
167:                        qualav = a2row(0).ToString
168:                    Else
169:                        qualav = 0
170:                    End If
171:                Case Is = 7
172:                    If Not IsNothing(a3row) Then
173:                        qualav = a3row(0).ToString
174:                    Else
175:                        qualav = 0
176:                    End If
177:                Case Is = 8
178:                    If Not IsNothing(a4row) Then
179:                        qualav = a4row(0).ToString
180:                    Else
181:                        qualav = 0
182:                    End If
183:                    End Select
184:        End If
185:        CodeEC1.Text = qualpresc
186:        codEC2.Text = qualav
187:        '4 era aqui
188:        If Not IsNothing(oque) Then
189:            Select Case oque
                    Case Is = 1
191:                    labelmeddci1.BackColor = Color.Red
192:                    labelmeddci.BackColor = Color.Red
193:                Case Is = 2
194:                    labelmedforma1.BackColor = Color.Red
195:                    labelmedforma.BackColor = Color.Red
196:                Case Is = 3
197:                    labelmeddose1.BackColor = Color.Red
198:                    labelmeddose.BackColor = Color.Red
199:                Case Is = 4
                        labelmedqty1.BackColor = Color.Red
201:                    labelmedqty.BackColor = Color.Red
200:                    labelmedcnpem1.BackColor = Color.Red
                        labelmedcnpem.BackColor = Color.Red
202:                Case Is = 5
203:                    labelmedcompgen1.BackColor = Color.Red
204:                    labelmedcompgen.BackColor = Color.Red
205:                Case Is = 6
206:
                        labelmedqty1.BackColor = Color.Yellow
207:                    labelmedqty.BackColor = Color.Yellow
208:                Case Is = 7
209:                    labelmedqty1.BackColor = Color.Red
                        labelmedqty.BackColor = Color.Red
                        labelmedcnpem1.BackColor = Color.Red
                        labelmedcnpem.BackColor = Color.Red
211:                    labelmedcompgen1.BackColor = Color.Orange
212:                    labelmedcompgen.BackColor = Color.Orange
213:                Case Is = 8
214:                    labelmedqty1.BackColor = Color.Yellow
215:                    labelmedqty.BackColor = Color.Yellow
216:                    labelmedcompgen1.BackColor = Color.Orange
217:                    labelmedcompgen.BackColor = Color.Orange
                    Case Is = 9
218:                    labelmedcnpem1.BackColor = Color.Red
219:                    labelmedcnpem.BackColor = Color.Red
220:
222:                Case Is = 10
223:                    labelmednome1.BackColor = Color.Red
224:                    labelmednome.BackColor = Color.Red
225:
226:                    End Select
227:        End If
228:    End If
229:    Exit Sub
MOSTRARERRO:
        MsgBox("SUB transferros: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub somas()
1:      On Error GoTo MOSTRARERRO
        If firsttime = False Then
a1:         Dim VVV = 0
2:          agrupar()
            For i = 1 To A
                indicar(i)
                indicar2(i)
            Next
3:          SomarPVP()
4:          SomarComp()
a4:         SomarPVP2()
a5:         SomarComp2()
5:          totalPVP.Text = SomarPVP()
6:          totalComp.Text = SomarComp()
a6:         totalPVP2.Text = SomarPVP2()
b6:         totalComp2.Text = SomarComp2()
c6:         'VVV = versehaerros()
7:          If (VVV >= 1 And VVV <= 799) Or VVV = 911 Or (VVV >= 955 And VVV <= 988) Or (VVV >= 855 And VVV <= 858) Then
a7:             vermelho = True
8:              totalComp.BackColor = Color.Red
a8:             totalComp2.BackColor = Color.Red
9:          ElseIf VVV > 801 And VVV <= 854 Then
a9:             amarelo = True
10:             totalComp.BackColor = Color.Orange
a10:            totalComp2.BackColor = Color.Orange
11:         ElseIf VVV = 0 Then
12:             totalComp.BackColor = Color.Green
a12:            totalComp2.BackColor = Color.Green
b12:        Else : MsgBox("versehaerros com valor desconhecido")
13:         End If
            totalComp.BackColor = verificador
            totalComp2.BackColor = verificador
            'If labeltroca1.Text <> "" Or labeltroca2.Text <> "" Or labeltroca3.Text <> "" Or labeltroca4.Text <> "" Or labelexcepab1.Text <> "" Or labelexcepab2.Text <> "" Or labelexcepab3.Text <> "" Or labelexcepab4.Text <> "" Then
            ' totalComp.BackColor = Color.Yellow
            ' totalComp2.BackColor = Color.Yellow
            ' End If
            If VVV <> 0 Then
14:             transferros(VVV)
15:         End If
        End If
16:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB somas: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'Private Sub but1474_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_01.CheckedChanged
    '1:      On Error GoTo MOSTRARERRO
    '        If Not IsNothing(a1row) Then
    '            If port1474 = True Then
    '2:              If but1474_01.Checked Then
    '3:                  portimedio = port1.Text
    '4:                  If a1row(17) = "true" And a1row(6) <> 0 Then
    '5:                      portcomp = Replace(0.69, ".", ",")
    '6:                  ElseIf a1row(18) = "true" And a1row(6) <> 0 Then
    '7:                      portcomp = 1
    '8:                  Else
    '9:                      portcomp = 0
    '10:                 End If
    '11:                 indicar(1)
    '12:                 somas()
    '13:             End If
    '            End If
    '        End If
    '14:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB but1474_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    '    Private Sub but1474_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_02.CheckedChanged
    '1:      On Error GoTo MOSTRARERRO
    '        If Not IsNothing(a2row) Then
    '            If port1474 = True Then
    '2:              If but1474_02.Checked Then
    '3:                  portimedio = port2.Text
    '4:                  If a2row(17) = "true" And a2row(6) <> 0 Then
    '5:                      portcomp = Replace(0.69, ".", ",")
    '6:                  ElseIf a2row(18) = "true" And a2row(6) <> 0 Then
    '7:                      portcomp = 1
    '8:                  Else
    '9:                      portcomp = 0
    '10:                 End If
    '11:                 indicar(2)
    '12:                 somas()
    '13:             End If
    '            End If
    '        End If
    '14:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB but1474_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    '    Private Sub but1474_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_03.CheckedChanged
    '1:      On Error GoTo MOSTRARERRO
    '        If Not IsNothing(a3row) Then
    '            If port1474 = True Then
    '2:              If but1474_03.Checked Then
    '3:                  portimedio = port3.Text
    '4:                  If a3row(17) = "true" And a3row(6) <> 0 Then
    '5:                      portcomp = Replace(0.69, ".", ",")
    '6:                  ElseIf a3row(18) = "true" And a3row(6) <> 0 Then
    '7:                      portcomp = 1
    '8:                  Else
    '9:                      portcomp = 0
    '10:                 End If
    '11:                 indicar(3)
    '12:                 somas()
    '13:             End If
    '       End If
    '   End If
    '14:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB but1474_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '  End Sub

    '    Private Sub but1474_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_04.CheckedChanged
    '1:      On Error GoTo MOSTRARERRO
    '        If Not IsNothing(a4row) Then
    '            If port1474 = True Then
    '2:              If but1474_04.Checked Then
    '3:                  If a4row(17) = "true" And a4row(6) <> 0 Then
    '4:                      portcomp = Replace(0.69, ".", ",")
    '5:                  ElseIf a4row(18) = "true" And a4row(6) <> 0 Then
    '6:                      portcomp = 1
    '7:                  Else
    '8:                      portcomp = 0
    '9:                  End If
    '10:                 indicar(4)
    '11:                 somas()
    '12:             End If
    '            End If
    '        End If
    '13:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB but1474_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    'Private Sub but1474_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_05.CheckedChanged
    '1:      On Error GoTo MOSTRARERRO
    '        If Not IsNothing(a5row) Then
    '            If port1474 = True Then
    '2:              If but1474_05.Checked Then
    '3:                  portimedio = port5.Text
    '4:                  If a5row(17) = "true" And a5row(6) <> 0 Then
    '5:                      portcomp = Replace(0.69, ".", ",")
    '6:                  ElseIf a5row(18) = "true" And a5row(6) <> 0 Then
    '7:                      portcomp = 1
    '8:                  Else
    '9:                      portcomp = 0
    '10:                 End If
    '11:                 indicar(5)
    '12:                 somas()
    '13:             End If
    '            End If
    '        End If
    '14:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB but1474_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub



    Private Sub but1234_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but1234_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(11) = "true" And a1row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(1)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but1234_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(11) = "true" And a2row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(2)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but1234_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(11) = "true" And a3row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(3)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but1234_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(11) = "true" And a4row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(4)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Private Sub but10279_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but10279_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(13) = "true" Or a1row(14) = "true" Then
5:                  If a1row(6) <> 0 Then
6:                      portcomp = Replace(tectocomp, ".", ",")
7:                  Else
8:                      portcomp = 0
9:                  End If
10:                 indicar(1)
11:                 somas()
12:             End If
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but10279_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(13) = "true" Or a2row(14) = "true" Then
5:                  If a2row(6) <> 0 Then
6:                      portcomp = Replace(tectocomp, ".", ",")
7:                  Else
8:                      portcomp = 0
9:                  End If
10:                 indicar(2)
11:                 somas()
12:             End If
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but10279_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(13) = "true" Or a3row(14) = "true" Then
5:                  If a3row(6) <> 0 Then
6:                      portcomp = Replace(tectocomp, ".", ",")
7:                  Else
8:                      portcomp = 0
9:                  End If
10:                 indicar(3)
11:                 somas()
12:             End If
            End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but10279_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(13) = "true" Or a4row(14) = "true" Then
5:                  If a4row(6) <> 0 Then
6:                      portcomp = Replace(tectocomp, ".", ",")
7:                  Else
8:                      portcomp = 0
9:                  End If
10:                 indicar(4)
11:                 somas()
12:             End If
            End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







    Private Sub but14123_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but14123_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(16) = "true" And a1row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(1)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but14123_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(16) = "true" And a2row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(2)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but14123_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(16) = "true" And a3row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(3)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but14123_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(16) = "true" And a4row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(4)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Private Sub but10910_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but10910_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(15) = "true" And a1row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(1)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but10910_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(15) = "true" And a2row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(2)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but10910_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(15) = "true" And a3row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(3)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but10910_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(15) = "true" And a4row(6) <> 0 Then
5:                  portcomp = Replace(0.69, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(4)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Private Sub but62010_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but62010_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but62010_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(19) = "true" And a1row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(1)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but62010_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but62010_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but62010_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but62010_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(19) = "true" And a2row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(2)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but62010_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but62010_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but62010_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but62010_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(19) = "true" And a3row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(3)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but62010_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but62010_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but62010_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but62010_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(19) = "true" And a4row(6) <> 0 Then
5:                  portcomp = Replace(tectocomp, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(4)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but62010_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Private Sub but21094_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but21094_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(12) = "true" And a1row(6) <> 0 Then
5:                  portcomp = 1
6:              Else
7:                  portcomp = 0
8:              End If
9:              For i = 1 To 6
10:                 indicar(i)
11:             Next
12:             somas()
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but21094_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(12) = "true" And a2row(6) <> 0 Then
5:                  portcomp = 1
6:              Else
7:                  portcomp = 0
8:              End If
9:              For i = 1 To 6
10:                 indicar(i)
11:             Next
12:             somas()
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but21094_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(12) = "true" And a3row(6) <> 0 Then
5:                  portcomp = 1
6:              Else
7:                  portcomp = 0
8:              End If
9:              For i = 1 To 6
10:                 indicar(i)
11:             Next
12:             somas()
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but21094_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but21094_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(12) = "true" And a4row(6) <> 0 Then
5:                  portcomp = 1
6:              Else
7:                  portcomp = 0
8:              End If
9:              For i = 1 To 6
10:                 indicar(i)
11:             Next
12:             somas()
13:         End If
        End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Private Sub but4250_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) Then
2:          If but4250_01.Checked Then
3:              portimedio = port1.Text
4:              If a1row(10) = "true" Then
5:                  portcomp = Replace(0.37, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(1)
10:             somas()
11:         End If
12:         If but42.Checked = True Then
13:             paramiloidose = True
14:         End If
15:         If but42.Checked = False Then
16:             paramiloidose = False
17:         End If
        End If
18:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) Then
2:          If but4250_02.Checked Then
3:              portimedio = port2.Text
4:              If a2row(10) = "true" Then
5:                  portcomp = Replace(0.37, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(2)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) Then
2:          If but4250_03.Checked Then
3:              portimedio = port3.Text
4:              If a3row(10) = "true" Then
5:                  portcomp = Replace(0.37, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(3)
10:             somas()
            End If
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) Then
2:          If but4250_04.Checked Then
3:              portimedio = port4.Text
4:              If a4row(10) = "true" Then
5:                  portcomp = Replace(0.37, ".", ",")
6:              Else
7:                  portcomp = 0
8:              End If
9:              indicar(4)
10:             somas()
11:         End If
        End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Private Sub but01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but01.Checked Then
3:          deslabelar()
            but01.Checked = True
4:          organismo = "01"
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but48_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but48.Checked Then
3:          deslabelar()
            but48.Checked = True
4:          organismo = 48
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but41_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but41.Checked Then
3:          deslabelar()
            but41.Checked = True
4:          organismo = 41
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "doentes profissionais"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but41_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but46_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but46.Checked Then
3:          deslabelar()
            but46.Checked = True
4:          organismo = 46
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but46_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but42_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but42.Checked Then
3:          deslabelar()
            but42.Checked = True
4:          organismo = 42
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "portaria 4521/2001"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but42_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but67_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but67.Checked Then
3:          deslabelar()
            but67.Checked = True
4:          organismo = 67
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "despacho 11387-A/2003"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but67_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub butDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If butDS.Checked Then
3:          deslabelar()
            butDS.Checked = True
4:          organismo = "ds"
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "diabetes SNS"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB butDS_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but49_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but49.Checked Then
3:          deslabelar()
            but49.Checked = True
4:          organismo = 49
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            If port1474 = True Then
7:              labelatribuido.Text = "portaria 1474/2004"
            End If
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but49_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub but45_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but45.Checked Then
3:          deslabelar()
            but45.Checked = True
4:          organismo = 45
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            If port1474 = True Then
7:              labelatribuido.Text = "portaria 1474/2004"
            End If
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but45_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub tirarports()
1:      On Error GoTo MOSTRARERRO
2:      but1474_01.Checked = False
3:      but1474_02.Checked = False
4:      but1474_03.Checked = False
5:      but1474_04.Checked = False
6:      'but1474_05.Checked = False
7:      'but1474_06.Checked = False
8:      but1234_01.Checked = False
9:      but1234_02.Checked = False
10:     but1234_03.Checked = False
11:     but1234_04.Checked = False
12:     'but1234_05.Checked = False
13:     'but1234_06.Checked = False
14:     but4250_01.Checked = False
15:     but4250_02.Checked = False
16:     but4250_03.Checked = False
17:     but4250_04.Checked = False
18:     'but4250_05.Checked = False
19:     'but4250_06.Checked = False
20:     but14123_01.Checked = False
21:     but14123_02.Checked = False
22:     but14123_03.Checked = False
23:     but14123_04.Checked = False
24:     'but14123_05.Checked = False
25:     'but14123_06.Checked = False
26:     but21094_01.Checked = False
27:     but21094_02.Checked = False
28:     but21094_03.Checked = False
29:     but21094_04.Checked = False
30:     'but21094_05.Checked = False
31:     'but21094_06.Checked = False
32:     but10279_01.Checked = False
33:     but10279_02.Checked = False
34:     but10279_03.Checked = False
35:     but10279_04.Checked = False
36:     'but10279_05.Checked = False
37:     'but10279_06.Checked = False
38:     but10910_01.Checked = False
39:     but10910_02.Checked = False
40:     but10910_03.Checked = False
41:     but10910_04.Checked = False
42:     'but10910_05.Checked = False
43:     'but10910_06.Checked = False
44:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB tirarports: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub deslabelar()
1:      On Error GoTo MOSTRARERRO
2:      labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Regular)
3:      labelatribuido.Text = ""
        desbutonizar()
4:      Exit Sub
MOSTRARERRO:
        MsgBox("SUB deslabelar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub organismus(ByVal umdeles As String)
1:      On Error GoTo MOSTRARERRO
2:      Select Case umdeles
            Case "01", 46, 41, 42, 67, 23
4:              tirarports()
5:          Case 49, 45, 59
                If port1474 = True Then
6:                  but1474_01.Checked = True
7:                  but1474_02.Checked = True
8:                  but1474_03.Checked = True
9:                  but1474_04.Checked = True
10:                 ' but1474_05.Checked = True
11:                 ' but1474_06.Checked = True
                End If
12:             End Select
        'organismotxt.Text = organismo
13:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB organismus: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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



    Function verifdiab(ByVal insub As Integer, ByVal qual As Short) As Short
1:      On Error GoTo MOSTRARERRO

2:      Dim resultado As Short
3:      resultado = 99
4:      If butDS.Checked = True Or butDM.Checked = True Then
5:          If diabetes(insub) = True Then
6:              resultado = 0
7:          Else
8:              resultado = 1
9:          End If
10:     Else
11:         If diabetes(insub) = True Then
12:             resultado = 2
13:         Else
14:             resultado = 3
15:         End If
16:     End If
17:
18:     If resultado = 1 Then
19:         vermelho = True
20:         Select Case qual
                Case Is = 1
22:                 conjunto = 11
24:             Case Is = 2
25:                 conjunto = 21
27:             Case Is = 3
28:                 conjunto = 31
30:             Case Is = 4
31:                 conjunto = 41
33:             Case Is = 5
34:                 conjunto = 51
36:             Case Is = 6
37:                 conjunto = 61
39:                 End Select
40:     End If
41:
42:     If resultado = 2 Then
43:         vermelho = True
44:         Select Case qual
                Case Is = 1
                    conjunto = 12
48:             Case Is = 2
                    conjunto = 22
51:             Case Is = 3
                    conjunto = 32
54:             Case Is = 4
55:                 conjunto = 42
57:             Case Is = 5
58:                 conjunto = 52
60:             Case Is = 6
61:                 conjunto = 62
63:                 End Select
64:     End If
65:
66:     'legenda
67:     '0 = diabetes em diabetes (ok)
68:     '1 = medicamento em diabetes
69:     '2 = diabetes em medicamento
70:     '3 = medicamento em medicamentos (ok)
71:     Exit Function
MOSTRARERRO:
        MsgBox("SUB verifdiab: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function


    Sub mostrardiab()
1:      On Error GoTo MOSTRARERRO
2:
3:      Dim what As Short
4:      Dim qualeh As Short
5:
6:      Select Case conjunto
            Case Is = 11
8:              what = 1
9:              qualeh = 1
10:         Case Is = 21
11:             what = 1
12:             qualeh = 2
13:         Case Is = 31
14:             what = 1
15:             qualeh = 3
16:         Case Is = 41
17:             what = 1
18:             qualeh = 4
19:             'Case Is = 51
20:             '   what = 1
21:             '  qualeh = 5
22:             'Case Is = 61
23:             '   what = 1
24:             '  qualeh = 6
25:         Case Is = 12
26:             what = 2
27:             qualeh = 1
28:         Case Is = 22
29:             what = 2
30:             qualeh = 2
31:         Case Is = 32
32:             what = 2
33:             qualeh = 3
34:         Case Is = 42
35:             what = 2
36:             qualeh = 4
37:             'Case Is = 52
38:             '   what = 2
39:             '  qualeh = 5
40:             'Case Is = 62
41:             '  what = 2
42:             '   qualeh = 6
43:             End Select
44:
45:
46:     If what = 1 Then
47:         vermelho = True
48:         Select Case qualeh
                Case Is = 1
49:                 result1.Text = "medicamento em lote de diabetes"
50:                 result1.BackColor = Color.Red
51:             Case Is = 2
52:                 result2.Text = "medicamento em lote de diabetes"
53:                 result2.BackColor = Color.Red
54:             Case Is = 3
55:                 result3.Text = "medicamento em lote de diabetes"
56:                 result3.BackColor = Color.Red
57:             Case Is = 4
58:                 result4.Text = "medicamento em lote de diabetes"
59:                 result4.BackColor = Color.Red
60:                 'Case Is = 5
61:                 '    result5.Text = "medicamento em lote de diabetes"
62:                 '    result5.BackColor = Color.Red
63:                 'Case Is = 6
64:                 '    result6.Text = "medicamento em lote de diabetes"
65:                 '    result6.BackColor = Color.Red
66:                 End Select
67:     End If
68:
69:     If what = 2 Then
70:         vermelho = True
71:         Select Case qualeh
                Case Is = 1
72:                 result1.Text = "diabetes em lote de medicamentos"
73:                 result1.BackColor = Color.Red
74:             Case Is = 2
75:                 result2.Text = "diabetes em lote de medicamentos"
76:                 result2.BackColor = Color.Red
77:             Case Is = 3
78:                 result3.Text = "diabetes em lote de medicamentos"
79:                 result3.BackColor = Color.Red
80:             Case Is = 4
81:                 result4.Text = "diabetes em lote de medicamentos"
82:                 result4.BackColor = Color.Red
83:                 '  Case Is = 5
84:                 '      result5.Text = "diabetes em lote de medicamentos"
85:                 '      result5.BackColor = Color.Red
86:                 '  Case Is = 6
87:                 '      result1.Text = "diabetes em lote de medicamentos"
88:                 '      result6.BackColor = Color.Red
89:                 End Select
90:     End If


        Exit Sub
MOSTRARERRO:
        MsgBox("SUB mostrardiab: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


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


    Private Sub TrocadelaboratorioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrocaDeLaboratórioToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      If filtrolab = False Then
            filtrolab = True
            TrocaDeLaboratórioToolStripMenuItem.Checked = True
        ElseIf filtrolab = True Then
            filtrolab = False
            TrocaDeLaboratórioToolStripMenuItem.Checked = False
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub MarcaMarcaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarcaMarcaToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      If filtromarcamarcadci = False Then
            filtromarcamarcadci = True
            MarcaMarcaToolStripMenuItem.Checked = True
        ElseIf filtromarcamarcadci = True Then
            filtromarcamarcadci = False
            MarcaMarcaToolStripMenuItem.Checked = False
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Private Sub AcederToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcederToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      form2.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub ECToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ECToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      EC.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Private Sub EPPRToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EPPRToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      EP.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub but01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but01.Click
        On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but01.Checked Then
3:          deslabelar()
            but01.Checked = True
4:          organismo = "01"
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but01_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but42_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but42.Checked Then
3:          deslabelar()
            but42.Checked = True
4:          organismo = 42
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but42_CheckChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but48.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but48.Checked Then
3:          deslabelar()
            but48.Checked = True
4:          organismo = 48
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but48_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but41.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but41.Checked Then
3:          deslabelar()
            but41.Checked = True
4:          organismo = 41
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "doentes profissionais"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but41_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but46.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but46.Checked Then
3:          deslabelar()
            but46.Checked = True
4:          organismo = 46
5:          organismus(organismo)
6:          aviam1.Focus()
7:      End If
8:      For i = 1 To 6
9:          indicar(i)
10:         somas()
11:     Next
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but46_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but42.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but42.Checked Then
3:          deslabelar()
            but42.Checked = True
4:          organismo = 42
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "portaria 4521/2001"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but42_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but67.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but67.Checked Then
3:          deslabelar()
            but67.Checked = True
4:          organismo = 67
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "despacho 11387-A/2003"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but67_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub butDS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDS.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If butDS.Checked Then
3:          deslabelar()
            butDS.Checked = True
4:          organismo = "DS"
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "diabetes SNS"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB butDS_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but49.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but49.Checked Then
3:          deslabelar()
            but49.Checked = True
4:          organismo = 49
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            If port1474 = True Then
7:              labelatribuido.Text = "portaria 1474/2004"
            End If
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but49_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but45.Click
1:      On Error GoTo MOSTRARERRO
        If limpo = True Then
            inicializar()
        End If
2:      If but45.Checked Then
3:          deslabelar()
            but45.Checked = True
4:          organismo = 45
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            If port1474 = True Then
7:              labelatribuido.Text = "portaria 1474/2004"
            End If
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but45_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub desbutonizar()
1:      On Error GoTo MOSTRARERRO
        but01.Checked = False
        'but03.Checked = False
        'but04.Checked = False
        'but06.Checked = False
        'but07.Checked = False
        'but08.Checked = False
        'but09.Checked = False
        'but10.Checked = False
        'but11.Checked = False
        'but14.Checked = False
        'but19.Checked = False
        'but20.Checked = False
        'but21.Checked = False
        'but22.Checked = False
        'but23.Checked = False
        'but24.Checked = False
        but25.Checked = False
        'but26.Checked = False
        'but27.Checked = False
        'but28.Checked = False
        'but29.Checked = False
        'but30.Checked = False
        'but31.Checked = False
        'but32.Checked = False
        'but33.Checked = False
        'but34.Checked = False
        'but35.Checked = False
        'but36.Checked = False
        'but37.Checked = False
        'but38.Checked = False
        'but39.Checked = False
        'but40.Checked = False
        but41.Checked = False
        but42.Checked = False
        'but43.Checked = False
        'but44.Checked = False
        but45.Checked = False
        but46.Checked = False
        but47.Checked = False
        but48.Checked = False
        but49.Checked = False
        'but50.Checked = False
        'but51.Checked = False
        'but52.Checked = False
        'but53.Checked = False
        'but54.Checked = False
        'but55.Checked = False
        'but56.Checked = False
        'but60.Checked = False
        'but61.Checked = False
        'but62.Checked = False
        'but63.Checked = False
        'but64.Checked = False
        'but65.Checked = False
        'but66.Checked = False
        but67.Checked = False
        'but69.Checked = False
        'but70.Checked = False
        'but71.Checked = False
        'but72.Checked = False
        'but73.Checked = False
        'but74.Checked = False
        'but75.Checked = False
        'but76.Checked = False
        'but77.Checked = False
        but85.Checked = False
        but86.Checked = False
        but87.Checked = False
        but88.Checked = False
        but89.Checked = False
        but90.Checked = False
        but91.Checked = False
        but92.Checked = False
        but93.Checked = False
        but94.Checked = False
        but95.Checked = False
        but96.Checked = False
        but97.Checked = False
        but98.Checked = False
        but99.Checked = False
        'butA1.Checked = False
        'butA2.Checked = False
        'butA3.Checked = False
        'butA4.Checked = False
        'butA5.Checked = False
        'butA6.Checked = False
        'butA7.Checked = False
        'butA8.Checked = False
        'butA9.Checked = False
        'butAA.Checked = False
        'butAB.Checked = False
        'butAC.Checked = False
        'butAD.Checked = False
        'butAE.Checked = False
        'butAF.Checked = False
        'butAG.Checked = False
        'butAH.Checked = False
        'butAI.Checked = False
        'butAJ.Checked = False
        'butB1.Checked = False
        'butB2.Checked = False
        'butB3.Checked = False
        'butB4.Checked = False
        'butB5.Checked = False
        'butB6.Checked = False
        'butB7.Checked = False
        'butB8.Checked = False
        'butBA.Checked = False
        'butBB.Checked = False
        'butBC.Checked = False
        'butBD.Checked = False
        'butBE.Checked = False
        'butBF.Checked = False
        'butBG.Checked = False
        'butBH.Checked = False
        'butBI.Checked = False
        'butBL.Checked = False
        'butBM.Checked = False
        'butBN.Checked = False
        'butBP.Checked = False
        'butBQ.Checked = False
        'butBR.Checked = False
        'butBS.Checked = False
        'butBT.Checked = False
        'butBU.Checked = False
        butBV.Checked = False
        'butBW.Checked = False
        butBX.Checked = False
        'butBY.Checked = False
        'butBZ.Checked = False
        'butC1.Checked = False
        'butC2.Checked = False
        'butC3.Checked = False
        'butC4.Checked = False
        'butC5.Checked = False
        'butC6.Checked = False
        'butC7.Checked = False
        'butC8.Checked = False
        'butCA.Checked = False
        'butCB.Checked = False
        'butCC.Checked = False
        'butCD.Checked = False
        'butCE.Checked = False
        'butCF.Checked = False
        'butCG.Checked = False
        'butCH.Checked = False
        'butCI.Checked = False
        'butCJ.Checked = False
        'butCL.Checked = False
        'butCM.Checked = False
        'butCN.Checked = False
        'butCO.Checked = False
        'butCP.Checked = False
        'butCQ.Checked = False
        'butCR.Checked = False
        'butCT.Checked = False
        'butCZ.Checked = False
        'butD1.Checked = False
        'butD2.Checked = False
        'butD3.Checked = False
        'butD4.Checked = False
        'butD5.Checked = False
        'butD6.Checked = False
        'butD7.Checked = False
        'butD8.Checked = False
        'butD9.Checked = False
        'butDA.Checked = False
        butDM.Checked = False
        butDS.Checked = False
        'butDT.Checked = False
        'butDU.Checked = False
        'butDX.Checked = False
        'butDY.Checked = False
        'butE1.Checked = False
        'butE2.Checked = False
        'butE3.Checked = False
        'butE4.Checked = False
        'butE5.Checked = False
        'butE6.Checked = False
        'butE7.Checked = False
        'butE8.Checked = False
        'butE9.Checked = False
        'butEA.Checked = False
        'butEB.Checked = False
        'butEC.Checked = False
        'butED.Checked = False
        'butEE.Checked = False
        'butEF.Checked = False
        'butEG.Checked = False
        'butEH.Checked = False
        'butEI.Checked = False
        'butEJ.Checked = False
        'butEL.Checked = False
        'butEM.Checked = False
        'butEN.Checked = False
        'butEO.Checked = False
        'butEP.Checked = False
        'butEQ.Checked = False
        'butF0.Checked = False
        butF1.Checked = False
        'butF2.Checked = False
        'butF4.Checked = False
        butF7.Checked = False
        'butF8.Checked = False
        'butF9.Checked = False
        'butFA.Checked = False
        'butFB.Checked = False
        'butFC.Checked = False
        butFM.Checked = False
        'butG1.Checked = False
        'butG2.Checked = False
        'butG3.Checked = False
        'butG4.Checked = False
        'butG5.Checked = False
        'butG6.Checked = False
        'butH1.Checked = False
        'butH3.Checked = False
        'butH4.Checked = False
        'butH5.Checked = False
        'butH6.Checked = False
        'butJ1.Checked = False
        'butJ2.Checked = False
        'butJ3.Checked = False
        'butJ4.Checked = False
        'butJ7.Checked = False
        'butJ8.Checked = False
        'butJ9.Checked = False
        'butJB.Checked = False
        'butL1.Checked = False
        'butL3.Checked = False
        'butL4.Checked = False
        'butL5.Checked = False
        'butL6.Checked = False
        'butL7.Checked = False
        'butLS.Checked = False
        'butM1.Checked = False
        'butM2.Checked = False
        'butM3.Checked = False
        'butM4.Checked = False
        'butM7.Checked = False
        'butM8.Checked = False
        'butN6.Checked = False
        'butO1.Checked = False
        'butO2.Checked = False
        'butO3.Checked = False
        'butO4.Checked = False
        'butO5.Checked = False
        'butO6.Checked = False
        'butO7.Checked = False
        'butO8.Checked = False
        'butO9.Checked = False
        'butOA.Checked = False
        butP1.Checked = False
        'butPA.Checked = False
        'butPB.Checked = False
        'butPC.Checked = False
        'butPD.Checked = False
        'butPR.Checked = False
        'butR1.Checked = False
        'butR2.Checked = False
        'butR3.Checked = False
        'butR4.Checked = False
        'butR5.Checked = False
        'butR6.Checked = False
        'butR7.Checked = False
        'butR8.Checked = False
        'butR9.Checked = False
        'butRA.Checked = False
        'butS1.Checked = False
        'butS2.Checked = False
        'butS3.Checked = False
        'butS4.Checked = False
        'butS5.Checked = False
        'butS6.Checked = False
        'butS7.Checked = False
        'butS8.Checked = False
        'butSA.Checked = False
        'butSB.Checked = False
        'butSC.Checked = False
        'butSP.Checked = False
        'butT1.Checked = False
        'butT2.Checked = False
        'butT3.Checked = False
        'butT4.Checked = False
        'butT5.Checked = False
        'butT6.Checked = False
        'butT7.Checked = False
        'butT8.Checked = False
        'butT9.Checked = False
        'butTA.Checked = False
        'butTB.Checked = False
        'butTC.Checked = False
        'butU1.Checked = False
        'butU2.Checked = False
        'butU3.Checked = False
        'butU4.Checked = False
        'butU5.Checked = False
        'butU6.Checked = False
        'butU7.Checked = False
        'butU8.Checked = False
        'butU9.Checked = False
        'butW0.Checked = False
        'butW1.Checked = False
        'butW2.Checked = False
        'butW3.Checked = False
        'butW4.Checked = False
        'butW5.Checked = False
        'butW6.Checked = False
        'butW7.Checked = False
        'butW8.Checked = False
        'butW9.Checked = False
        'butWA.Checked = False
        'butX0.Checked = False
        'butX1.Checked = False
        'butX2.Checked = False
        'butX3.Checked = False
        'butX4.Checked = False
        'butX5.Checked = False
        'butX6.Checked = False
        'butX7.Checked = False
        'butX8.Checked = False
        'butX9.Checked = False
        'butXA.Checked = False
        'butXB.Checked = False
        'butXC.Checked = False
        'butXD.Checked = False
        'butXE.Checked = False
        'butXF.Checked = False
        'butXG.Checked = False
        'butXH.Checked = False
        'butXI.Checked = False
        'butXJ.Checked = False
        'butXK.Checked = False
        'butXL.Checked = False
        'butXM.Checked = False
        'butXN.Checked = False
        'butXO.Checked = False
        'butXP.Checked = False
        'butXQ.Checked = False
        'butXR.Checked = False
        'butXS.Checked = False
        'butXT.Checked = False
        'butXU.Checked = False
        'butXV.Checked = False
        'butXW.Checked = False
        'butXX.Checked = False
        'butXY.Checked = False
        'butXZ.Checked = False
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB desbutonizar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    '  Sub desfarmaciar()
    '1:      On Error GoTo MOSTRARERRO
    '       labelcodefarmacia.Text = ""
    '       labelNomeFarmacia.Text = ""
    '       b08168.Checked = False
    '       b04243.Checked = False
    '       b01813.Checked = False
    '       b03441.Checked = False
    '       b14400.Checked = False
    '       b03948.Checked = False
    '       b19542.Checked = False
    '       Exit Sub
    'MOSTRARERRO:
    '       MsgBox("SUB desfarmaciar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '   End Sub



    '  Sub farmaciar(ByVal qual As Integer)
    '1:      On Error GoTo MOSTRARERRO
    '2:      labelcodefarmacia.Text = qual
    '3:      Select Case qual
    '            Case Is = 8168
    '5:              labelNomeFarmacia.Text = "Central SJM"
    '6:          Case Is = 4243
    '7:              labelNomeFarmacia.Text = "Bessa OAZ"
    '8:          Case Is = 3441
    '9:              labelNomeFarmacia.Text = "Moderna Esmoriz"
    '10:         Case Is = 1813
    '11:             labelNomeFarmacia.Text = "Estação NINE"
    '12:         Case Is = 3948
    '13:             labelNomeFarmacia.Text = "Saraiva Avintes"
    '14:         Case Is = 14400
    '15:             labelNomeFarmacia.Text = "Aliança Vermoim"
    '            Case Is = 19542
    '                labelNomeFarmacia.Text = "Elvira Coelho"
    '            Case Is = 20575
    '               labelNomeFarmacia.Text = "Almeida e Sousa"
    '            Case Is = 13960
    '              labelNomeFarmacia.Text = "Araújo"
    '           Case Is = 25194
    '               labelNomeFarmacia.Text = "Santa Quitéria"
    '16:             End Select
    'MOSTRARERRO:
    '       MsgBox("SUB farmaciar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub

    '   Private Sub b08168_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b08168.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b08168.Checked = False Then
    '3:      End If
    '        If b08168.Checked = True Then
    '            farmacia = 8168
    '            desfarmaciar()
    '            MsgBox("5")
    '            farmaciar(farmacia)
    '        End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB b08168_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub
    '
    '    Private Sub b19542_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b19542.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b19542.Checked = False Then
    '3:      End If
    '       If b19542.Checked = True Then
    '           farmacia = 19542
    '           desfarmaciar()
    '           MsgBox("5")
    '          farmaciar(farmacia)
    '      End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '     MsgBox("SUB b19542_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '   End Sub

    ' Private Sub b04243_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b04243.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b04243.Checked = False Then
    '3:      End If
    '        If b04243.Checked = True Then
    '            farmacia = 4243
    '            desfarmaciar()
    '            farmaciar(farmacia)
    '        End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB b04243_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '   End Sub
    '
    '
    '    Private Sub b03441_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b03441.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b03441.Checked = False Then
    '3:      End If
    '       If b03441.Checked = True Then
    '           farmacia = 3441
    '           desfarmaciar()
    '           farmaciar(farmacia)
    '       End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '       MsgBox("SUB b03441_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '   End Sub
    '
    '    Private Sub b01813_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b01813.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b01813.Checked = False Then
    '3:      End If
    '        If b01813.Checked = True Then
    '            farmacia = 1813
    '            desfarmaciar()
    '            farmaciar(farmacia)
    '        End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB b01813_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '   End Sub
    '

    '    Private Sub b03948_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b03948.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b03948.Checked = False Then
    '3:      End If
    '        If b03948.Checked = True Then
    '            farmacia = 3948
    '            desfarmaciar()
    '            farmaciar(farmacia)
    '        End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB b03948_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    '    Private Sub b14400_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b14400.Click
    '1:      On Error GoTo MOSTRARERRO
    '2:      If b14400.Checked = False Then
    '3:      End If
    '       If b14400.Checked = True Then
    '           farmacia = 14400
    '           desfarmaciar()
    '           farmaciar(farmacia)
    '       End If
    '10:     Exit Sub
    'MOSTRARERRO:
    '        MsgBox("SUB b14400_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub



    '    Private Sub AbrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AbrirToolStripMenuItem.Click
    '41:     desfarmaciar()
    '    Dim janela As New abrir
    '        janela.Show()
    '        Me.Close()
    '    End Sub

    Private Sub RepovoarnovosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RepovoarnovosToolStripMenuItem.Click
        repovoar.Show()
    End Sub






    Sub subviadif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
        'obsoleto: comaparação da via é agora absoluto
        amarelo = True
        Select Case qual
            Case Is = 1
                '     labelsub1.Text = "S"
                '    labelsub1.BackColor = Color.Orange
                '   result1.BackColor = Color.BlueViolet
            Case Is = 2
                '     labelsub2.Text = "S"
                '     labelsub2.BackColor = Color.Orange
                '     result2.BackColor = Color.BlueViolet
            Case Is = 3
                '     labelsub3.Text = "S"
                '     labelsub3.BackColor = Color.Orange
                '    result3.BackColor = Color.BlueViolet
            Case Is = 4
                ' labelsub4.Text = "S"
                'labelsub4.BackColor = Color.Orange
                'result4.BackColor = Color.BlueViolet
                '   Case Is = 5
                '       labelsub5.Text = "S"
                '       labelsub5.BackColor = Color.Orange
                '      result5.BackColor = Color.BlueViolet
                '  Case Is = 6
                '      labelsub6.Text = "S"
                '      labelsub6.BackColor = Color.Orange
                '      result6.BackColor = Color.BlueViolet
        End Select

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub subviaDif: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub

    Sub marcadif(ByVal qual As Short, ByVal h As Short)
        On Error GoTo MOSTRARERRO
        'qual = 7 implica sem obrigatoriedade de prescrição por dci
        If filtrolab = True Or filtromarcamarcadci = True Then
            amarelo = True
            Select Case h
                Case Is = 0
                    'qty=
                    Select Case qual
                        Case Is = 1
                            '    labellab1.Text = "MARCA"
                            '    labellab1.BackColor = Color.Orange
                            result1.BackColor = Color.Orange
                            result1.Text = "marca diferente sem haver genéricos"
                        Case Is = 2
                            '    labellab2.Text = "MARCA"
                            '    labellab2.BackColor = Color.Orange
                            result2.BackColor = Color.Orange
                            result2.Text = "marca diferente sem haver genéricos"
                        Case Is = 3
                            '   labellab3.Text = "MARCA"
                            '   labellab3.BackColor = Color.Orange
                            result3.BackColor = Color.Orange
                            result3.Text = "marca diferente sem haver genéricos"
                        Case Is = 4
                            '  labellab4.Text = "MARCA"
                            '  labellab4.BackColor = Color.Orange
                            result4.BackColor = Color.Orange
                            result4.Text = "marca diferente sem haver genéricos"
                        Case Is = 7
                            ' labelsub1.Text = "nDCI"
                            ' labelsub1.BackColor = Color.Red
                            'result1.BackColor = Color.Orange
                            'result1.Text = "marca diferente sem haver genéricos"
                        Case Is = 8
                            ' labelsub2.Text = "nDCI"
                            ' labelsub2.BackColor = Color.Red
                            'result2.BackColor = Color.Orange
                            'result2.Text = "marca diferente sem haver genéricos"
                        Case Is = 9
                            ' labelsub3.Text = "nDCI"
                            ' labelsub3.BackColor = Color.Red
                            'result3.BackColor = Color.Orange
                            'result3.Text = "marca diferente sem haver genéricos"
                        Case Is = 10
                            ' labelsub4.Text = "nDCI"
                            ' labelsub4.BackColor = Color.Red
                            'result4.BackColor = Color.Orange
                            'result4.Text = "marca diferente sem haver genéricos"
                            '    Case Is = 5
                            '        labellab5.Text = "MARCA"
                            '        labellab5.BackColor = Color.Orange
                            '        result5.BackColor = Color.Orange
                            '        result1.Text = "marca diferente sem haver genéricos"
                            '    Case Is = 6
                            '       labellab6.Text = "MARCA"
                            '       labellab6.BackColor = Color.Orange
                            '       result6.BackColor = Color.Orange
                            '       result1.Text = "marca diferente sem haver genéricos"
                    End Select
                Case Is = 1
                    vermelho = True
                    Select Case qual
                        Case Is = 1
                            '   labellab1.Text = "MARCA"
                            '   labellab1.BackColor = Color.Orange
                            '   labelf1.Text = "h"
                            '   labelf1.BackColor = Color.BlueViolet
                            result1.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            result1.BackColor = Color.Red
                            transferros(411)
                        Case Is = 2
                            '  labellab2.Text = "MARCA"
                            '  labellab2.BackColor = Color.Orange
                            '  labelf2.Text = "h"
                            '  labelf2.BackColor = Color.BlueViolet
                            result2.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            result2.BackColor = Color.Red
                            transferros(412)
                        Case Is = 3
                            '  labellab3.Text = "MARCA"
                            '  labellab3.BackColor = Color.Orange
                            '  labelf3.Text = "h"
                            ' labelf3.BackColor = Color.BlueViolet
                            result3.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            result3.BackColor = Color.Red
                            transferros(413)
                        Case Is = 4
                            '  labellab4.Text = "MARCA"
                            '  labellab4.BackColor = Color.Orange
                            ' labelf4.Text = "h"
                            ' labelf4.BackColor = Color.BlueViolet
                            result4.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            result4.BackColor = Color.Red
                            transferros(414)
                        Case Is = 7
                            '  labelsub1.Text = "nDCI"
                            ' labelsub1.BackColor = Color.Red
                            ' labelf1.Text = "h"
                            ' labelf1.BackColor = Color.BlueViolet
                            'result1.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            'result1.BackColor = Color.Red
                            transferros(11)
                        Case Is = 8
                            ' labelsub2.Text = "nDCI"
                            ' labelsub2.BackColor = Color.Red
                            ' labelf2.Text = "h"
                            ' labelf2.BackColor = Color.BlueViolet
                            'result2.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            'result2.BackColor = Color.Red
                            transferros(12)
                        Case Is = 9
                            ' labelsub3.Text = "nDCI"
                            'labelsub3.BackColor = Color.Red
                            ' labelf3.Text = "h"
                            ' labelf3.BackColor = Color.BlueViolet
                            'result3.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            'result3.BackColor = Color.Red
                            transferros(13)
                        Case Is = 10
                            '  labelsub4.Text = "nDCI"
                            '  labelsub4.BackColor = Color.Red
                            ' labelf4.Text = "h"
                            ' labelf4.BackColor = Color.BlueViolet
                            'result4.Text = "h) h) h) h) e marca diferente sem haver genéricos"
                            'result4.BackColor = Color.Red
                            transferros(14)



                            '      Case Is = 5
                            '          labellab5.Text = "MARCA"
                            '          labellab5.BackColor = Color.Orange
                            '          labelf5.Text = "h"
                            '          labelf5.BackColor = Color.BlueViolet
                            '          result5.Text = "h) qty > 150%   ou  CNPEM"
                            '          result5.BackColor = Color.Red
                            '          MsgBox("final h=1")
                            '      Case Is = 6
                            '          labellab6.Text = "MARCA"
                            '          labellab6.BackColor = Color.Orange
                            '          labelf6.Text = "h"
                            '          labelf6.BackColor = Color.BlueViolet
                            '          result6.Text = "h) qty > 150%   ou  CNPEM"
                            '          result6.BackColor = Color.Red
                            '          MsgBox("final h=1")
                    End Select
                Case Is = 2
                    'qty<
                    Select Case qual
                        Case Is = 1
                            '          labellab1.Text = "MARCA"
                            '          labellab1.BackColor = Color.Orange
                            result1.BackColor = Color.Orange
                            result1.Text = "marca diferente sem haver genéricos"
                        Case Is = 2
                            '          labellab2.Text = "MARCA"
                            '          labellab2.BackColor = Color.Orange
                            result2.BackColor = Color.Orange
                            result2.Text = "marca diferente sem haver genéricos"
                        Case Is = 3
                            '         labellab3.Text = "MARCA"
                            '         labellab3.BackColor = Color.Orange
                            result3.BackColor = Color.Orange
                            result3.Text = "marca diferente sem haver genéricos"
                        Case Is = 4
                            '        labellab4.Text = "MARCA"
                            '        labellab4.BackColor = Color.Orange
                            result4.BackColor = Color.Orange
                            result4.Text = "marca diferente sem haver genéricos"
                            '   Case Is = 5
                            '       labellab5.Text = "MARCA"
                            '       labellab5.BackColor = Color.Orange
                            '       result5.BackColor = Color.Orange
                            '       result5.Text = "marca diferente"
                            '   Case Is = 6
                            '       labellab6.Text = "MARCA"
                            '       labellab6.BackColor = Color.Orange
                            '       result6.BackColor = Color.Orange
                            '       result6.Text = "marca diferente"
                    End Select
                Case Is = 3
                    'qty não numérica
                    Select Case qual
                        Case Is = 1
                            '        labellab1.Text = "MARCA"
                            '        labellab1.BackColor = Color.Orange
                            '     labelf1.Text = "h?"
                            '     labelf1.BackColor = Color.BlueViolet
                            result1.BackColor = Color.Orange
                            result1.Text = "VERIFQUANT e marca diferente sem haver genéricos"
                            transferros(611)
                        Case Is = 2
                            '        labellab2.Text = "MARCA"
                            '     labelf2.Text = "h?"
                            '     labelf2.BackColor = Color.BlueViolet
                            '        labellab2.BackColor = Color.Orange
                            result2.BackColor = Color.Orange
                            result2.Text = "VERIFQUANT e marca diferente sem haver genéricos"
                            transferros(612)
                        Case Is = 3
                            '       labellab3.Text = "MARCA"
                            '       labellab3.BackColor = Color.Orange
                            '    labelf3.Text = "h?"
                            '    labelf3.BackColor = Color.BlueViolet
                            result3.BackColor = Color.Orange
                            result3.Text = "VERIFQUANT e marca diferente sem haver genéricos"
                            transferros(613)
                        Case Is = 4
                            '      labellab4.Text = "MARCA"
                            '      labellab4.BackColor = Color.Orange
                            '   labelf4.Text = "h?"
                            '   labelf4.BackColor = Color.BlueViolet
                            result4.BackColor = Color.Orange
                            result4.Text = "VERIFQUANT e marca diferente sem haver genéricos"
                            transferros(614)
                            '   Case Is = 5
                            '       labellab5.Text = "MARCA"
                            '       labellab5.BackColor = Color.Orange
                            '       labelf5.Text = "h?"
                            '       labelf5.BackColor = Color.BlueViolet
                            '       result5.BackColor = Color.Orange
                            '   Case Is = 6
                            '       labellab6.Text = "MARCA"
                            '       labellab6.BackColor = Color.Orange
                            '       labelf6.Text = "h?"
                            '       labelf6.BackColor = Color.BlueViolet
                            '       result6.BackColor = Color.Orange
                    End Select
            End Select

        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub marcaDif: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub trocalab(ByVal qual As Short, ByVal de As Array, ByVal para As Array, ByVal h As Short)
        On Error GoTo MOSTRARERRO
        If filtrolab = True Then
            Select Case h
                Case Is = 0 'qty =
                    Select Case qual
                        Case Is = 1
                            amarelo = True
                            result1.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result1.BackColor = Color.Yellow
                            '                  labellab1.Text = "L"
                            '                  labellab1.BackColor = Color.BlueViolet
                        Case Is = 2
                            amarelo = True
                            result2.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result2.BackColor = Color.Yellow
                            '                  labellab2.Text = "L"
                            '                  labellab2.BackColor = Color.BlueViolet
                        Case Is = 3
                            amarelo = True
                            result3.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result3.BackColor = Color.Yellow
                            '                 labellab3.Text = "L"
                            '                 labellab3.BackColor = Color.BlueViolet
                        Case Is = 4
                            amarelo = True
                            result4.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result4.BackColor = Color.Yellow
                            '                labellab4.Text = "L"
                            '                labellab4.BackColor = Color.BlueViolet
                            '   Case Is = 5
                            '       amarelo = True
                            '       result5.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '       result5.BackColor = Color.Yellow
                            '       labellab5.Text = "L"
                            '       labellab5.BackColor = Color.BlueViolet
                            '   Case Is = 6
                            '       amarelo = True
                            '       result6.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '       result6.BackColor = Color.Yellow
                            '       labellab6.Text = "L"
                            '       labellab6.BackColor = Color.BlueViolet
                    End Select
                Case Is = 1 'qty >
                    Select Case qual
                        Case Is = 1
                            vermelho = True
                            result1.Text = "h) + verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result1.BackColor = Color.Red
                            '               labellab1.Text = "L"
                            '               labellab1.BackColor = Color.BlueViolet
                        Case Is = 2
                            vermelho = True
                            result2.Text = "h) + verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result2.BackColor = Color.Red
                            '               labellab2.Text = "L"
                            '               labellab2.BackColor = Color.BlueViolet
                        Case Is = 3
                            vermelho = True
                            result3.Text = "h) + verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result3.BackColor = Color.Red
                            '              labellab3.Text = "L"
                            '              labellab3.BackColor = Color.BlueViolet
                        Case Is = 4
                            vermelho = True
                            result4.Text = "h) + verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result4.BackColor = Color.Red
                            '             labellab4.Text = "L"
                            '             labellab4.BackColor = Color.BlueViolet
                            '   Case Is = 5
                            '       vermelho = True
                            '       result5.Text = "h) + verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '       result5.BackColor = Color.Red
                            '       labellab5.Text = "L"
                            '       labellab5.BackColor = Color.BlueViolet
                            '   Case Is = 6
                            '       vermelho = True
                            '       result6.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '       result6.BackColor = Color.Red
                            '       labellab6.Text = "L"
                            '       labellab6.BackColor = Color.BlueViolet
                    End Select

                Case Is = 2 'qty <=
                    Select Case qual
                        Case Is = 1
                            amarelo = True
                            result1.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result1.BackColor = Color.Yellow
                            '            labellab1.Text = "L"
                            '            labellab1.BackColor = Color.BlueViolet
                        Case Is = 2
                            amarelo = True
                            result2.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result2.BackColor = Color.Yellow
                            '           labellab2.Text = "L"
                            '           labellab2.BackColor = Color.BlueViolet
                        Case Is = 3
                            amarelo = True
                            result3.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result3.BackColor = Color.Yellow
                            '          labellab3.Text = "L"
                            '          labellab3.BackColor = Color.BlueViolet
                        Case Is = 4
                            amarelo = True
                            result4.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            result4.BackColor = Color.Yellow
                            '         labellab4.Text = "L"
                            '         labellab4.BackColor = Color.BlueViolet
                            '  Case Is = 5
                            '      amarelo = True
                            '      result5.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '      result5.BackColor = Color.Yellow
                            '      labellab5.Text = "L"
                            '      labellab5.BackColor = Color.BlueViolet
                            '  Case Is = 6
                            '      amarelo = True
                            '      result6.Text = "verificar desautorização (de " & de(8) & " para " & para(8) & ")"
                            '      result6.BackColor = Color.Yellow
                            '      labellab6.Text = "L"
                            '      labellab6.BackColor = Color.BlueViolet
                    End Select

                Case Is = 3
                    Select Case qual
                        Case Is = 1
                            amarelo = True
                            verifQuant(1)
                            '        labellab1.Text = "L"
                            '        labellab1.BackColor = Color.BlueViolet
                        Case Is = 2
                            amarelo = True
                            verifQuant(2)
                            '       labellab2.Text = "L"
                            '       labellab2.BackColor = Color.BlueViolet
                        Case Is = 3
                            amarelo = True
                            verifQuant(3)
                            '      labellab3.Text = "L"
                            '      labellab3.BackColor = Color.BlueViolet
                        Case Is = 4
                            amarelo = True
                            verifQuant(4)
                            '     labellab4.Text = "L"
                            '     labellab4.BackColor = Color.BlueViolet
                            ' Case Is = 5
                            '     amarelo = True
                            '     verifQuant(5)
                            '     labellab5.Text = "L"
                            '     labellab5.BackColor = Color.BlueViolet
                            ' Case Is = 6
                            '     amarelo = True
                            '     verifQuant(6)
                            '     labellab6.Text = "L"
                            '     labellab6.BackColor = Color.BlueViolet
                    End Select

            End Select


        End If

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub trocalab: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Function mesmonomeoulab(ByVal nomeprescrito As String, ByVal nomeaviado As String, ByVal labprescrito As String, labaviado As String)
        On Error GoTo MOSTRARERRO
        mesmonomeoulab = False
        If nomeprescrito = nomeaviado Or labprescrito = labaviado Then
            mesmonomeoulab = True
        End If
        Exit Function
MOSTRARERRO:
        MsgBox("Function mesmonomeoulab: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Private Sub NoCodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoCodeToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
1:      FormNoCode.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("sub NoCodeToolStripMenuItem: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'Sub CCF()
    'é temporário - até 31/3/2013
    '1:      On Error GoTo MOSTRARERRO
    '1000:   'MsgBox(p2row(18))
    '2:      Dim nDCI As String = "nDCI "
    '3:      Dim CCF1 As String = ""
    '4:      Dim CCF2 As String = ""
    '5:      Dim CCF3 As String = ""
    '6:      Dim CCF4 As String = ""
    '7:      Dim preCCF1 As String = ""
    '8:      Dim preCCF2 As String = ""
    '9:      Dim preCCF3 As String = ""
    '10:     Dim preCCF4 As String = ""
    '11:     preCCF1 = result1.Text
    '12:     preCCF2 = result2.Text
    '13:     preCCF3 = result3.Text
    '14:     preCCF4 = result4.Text
    '        Select Case P
    '            Case Is = 1
    '                Exit Sub
    '            Case Is = 2
    '                GoTo p2
    '            Case Is = 3
    '                GoTo p3
    '            Case Is = 4
    '                GoTo p4
    '        End Select
    '
    'p2:     If Not IsNothing(p2array) Then
    '16:         If p2row(18) = False Then
    '17:             If labelcruz1.Text = "[2] -> [1]" Then
    '18:                 If p2row(9).ToString <> a1row(9).ToString Then
    '19:                     CCF1 = nDCI + preCCF1
    '20:                     result1.Text = CCF1
    '21:                     result1.BackColor = Color.Red
    '22:                 End If
    '23:             End If
    '24:             If labelcruz2.Text = "[2] -> [2]" Then
    '25:                 If p2row(9).ToString <> a2row(9).ToString Then
    '26:                     CCF2 = nDCI + preCCF2
    '27:                     result2.Text = CCF2
    '28:                     result2.BackColor = Color.Red
    '29:                 End If
    '30:             End If
    '31:             If labelcruz3.Text = "[2] -> [3]" Then
    '32:                 If p2row(9).ToString <> a3row(9).ToString Then
    '33:                     CCF3 = nDCI + preCCF3
    '34:                     result3.Text = CCF3
    '35:                     result3.BackColor = Color.Red
    '36:                 End If
    '37:             End If
    '38:             If labelcruz4.Text = "[2] -> [4]" Then
    '39:                 If p2row(9).ToString <> a4row(9).ToString Then
    '40:                     CCF4 = nDCI + preCCF4
    '41:                     result4.Text = CCF4
    '42:                     result4.BackColor = Color.Red
    '43:                 End If
    '44:             End If
    '45:         End If
    '        End If
    '        Exit Sub
    'p3:     If Not IsNothing(p3array) Then
    '48:         If p3row(18) = False Then
    '49:             If labelcruz1.Text = "[3] -> [1]" Then
    '50:                 If p3row(9).ToString <> a1row(9).ToString Then
    '51:                     CCF1 = nDCI + preCCF1
    '52:                     result1.Text = CCF1
    '53:                     result1.BackColor = Color.Red
    '54:                 End If
    '55:             End If
    '56:             If labelcruz2.Text = "[3] -> [2]" Then
    '57:                 If p3row(9).ToString <> a2row(9).ToString Then
    '58:                     CCF2 = nDCI + preCCF2
    '59:                     result2.Text = CCF2
    '60:                     result2.BackColor = Color.Red
    '61:                 End If
    '62:             End If
    '63:             If labelcruz3.Text = "[3] -> [3]" Then
    '64:                 If p3row(9).ToString <> a3row(9).ToString Then
    '65:                     CCF3 = nDCI + preCCF3
    '66:                     result3.Text = CCF3
    '67:                     result3.BackColor = Color.Red
    '68:                 End If
    '69:             End If
    '70:             If labelcruz4.Text = "[3] -> [4]" Then
    '71:                 If p3row(9).ToString <> a4row(9).ToString Then
    '72:                     CCF4 = nDCI + preCCF4
    '73:                     result4.Text = CCF4
    '74:                     result4.BackColor = Color.Red
    '75:                 End If
    '76:             End If
    '77:         End If
    '        End If
    '        Exit Sub
    'p4:     If Not IsNothing(p4array) Then
    '80:         If p4row(18) = False Then
    '81:             If labelcruz1.Text = "[4] -> [1]" Then
    '82:                 If p4row(9).ToString <> a1row(9).ToString Then
    '83:                     CCF1 = nDCI + preCCF1
    '84:                     result1.Text = CCF1
    '85:                     result1.BackColor = Color.Red
    '86:                 End If
    '87:             End If
    '88:             If labelcruz2.Text = "[4] -> [2]" Then
    '89:                 If p4row(9).ToString <> a2row(9).ToString Then
    '90:                     CCF2 = nDCI + preCCF2
    '91:                     result2.Text = CCF2
    '92:                     result2.BackColor = Color.Red
    '93:                 End If
    '94:             End If
    '95:             If labelcruz3.Text = "[4] -> [3]" Then
    '96:                 If p4row(9).ToString <> a3row(9).ToString Then
    '97:                     CCF3 = nDCI + preCCF3
    '98:                     result3.Text = CCF3
    '99:                     result3.BackColor = Color.Red
    '100:                End If
    '101:            End If
    '102:            If labelcruz4.Text = "[4] -> [4]" Then
    '103:                If p4row(9).ToString <> a4row(9).ToString Then
    '104:                    CCF4 = nDCI + preCCF4
    '105:                    result4.Text = CCF4
    '106:                    result4.BackColor = Color.Red
    '107:                End If
    '108:            End If
    '109:        End If
    '110:    End If
    '111:    Exit Sub
    'MOSTRARERRO:
    '        MsgBox("function CCF: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '   End Sub






    Sub limparnofim()
        On Error GoTo MOSTRARERRO
        If aviam4.Text = "" Or aviam4.Text = "0" Then
            result4.Text = ""
            result4.BackColor = SystemColors.GradientInactiveCaption
        End If
        If aviam3.Text = "" Or aviam3.Text = "0" Then
            result3.Text = ""
            result3.BackColor = SystemColors.GradientInactiveCaption
        End If
        If aviam2.Text = "" Or aviam2.Text = "0" Then
            result2.Text = ""
            result2.BackColor = SystemColors.GradientInactiveCaption
        End If
        Select Case organismo
            Case Is = "ds"
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result1.Text = "medicamento em lote de diabetes"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam2.Text < 6190000 Or aviam2.Text > 6790000 Then
                    result2.Text = "medicamento em lote de diabetes"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text < 6190000 Or aviam3.Text > 6790000 Then
                    result3.Text = "medicamento em lote de diabetes"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result4.Text = "medicamento em lote de diabetes"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
            Case Is = "dj"
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result1.Text = "medicamento em lote de diabetes"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam2.Text < 6190000 Or aviam2.Text > 6790000 Then
                    result2.Text = "medicamento em lote de diabetes"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text < 6190000 Or aviam3.Text > 6790000 Then
                    result3.Text = "medicamento em lote de diabetes"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result4.Text = "medicamento em lote de diabetes"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
            Case Is = "dq"
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result1.Text = "medicamento em lote de diabetes"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam2.Text < 6190000 Or aviam2.Text > 6790000 Then
                    result2.Text = "medicamento em lote de diabetes"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text < 6190000 Or aviam3.Text > 6790000 Then
                    result3.Text = "medicamento em lote de diabetes"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result4.Text = "medicamento em lote de diabetes"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
            Case Is = "dp"
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result1.Text = "medicamento em lote de diabetes"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam2.Text < 6190000 Or aviam2.Text > 6790000 Then
                    result2.Text = "medicamento em lote de diabetes"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text < 6190000 Or aviam3.Text > 6790000 Then
                    result3.Text = "medicamento em lote de diabetes"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result4.Text = "medicamento em lote de diabetes"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
            Case Is = "dr"
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result1.Text = "medicamento em lote de diabetes"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam2.Text < 6190000 Or aviam2.Text > 6790000 Then
                    result2.Text = "medicamento em lote de diabetes"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text < 6190000 Or aviam3.Text > 6790000 Then
                    result3.Text = "medicamento em lote de diabetes"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam1.Text < 6190000 Or aviam1.Text > 6790000 Then
                    result4.Text = "medicamento em lote de diabetes"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
            Case Else
                If aviam1.Text > 6190000 And aviam1.Text < 6799999 Then
                    result1.Text = "diabetes em lote de medicamentos"
                    result1.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam4.Text > 6190000 And aviam4.Text < 6799999 Then
                    result2.Text = "diabetes em lote de medicamentos"
                    result2.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam3.Text > 6190000 And aviam3.Text < 6799999 Then
                    result3.Text = "diabetes em lote de medicamentos"
                    result3.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If
                If aviam4.Text > 6190000 And aviam4.Text < 6799999 Then
                    result4.Text = "diabetes em lote de medicamentos"
                    result4.BackColor = Color.Red
                    totalComp.BackColor = Color.Red
                End If

        End Select

        Exit Sub
MOSTRARERRO:
        MsgBox("sub limparnofim: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub traduzirnofim()
        On Error GoTo MOSTRARERRO


        Exit Sub

MOSTRARERRO:
        MsgBox("sub traduzirnofim: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'chama o irbuscar() e o row2array para produzir os resultados
    Private Sub Comparar()
        On Error GoTo MOSTRARERRO
        If comparado = False Then
1:          Me.result1.Text = ""
2:          Me.result1.BackColor = SystemColors.GradientInactiveCaption
3:          Me.result2.Text = ""
4:          Me.result2.BackColor = SystemColors.GradientInactiveCaption
5:          Me.result3.Text = ""
6:          Me.result3.BackColor = SystemColors.GradientInactiveCaption
7:          Me.result4.Text = ""
8:          Me.result4.BackColor = SystemColors.GradientInactiveCaption

9:          irbuscar2013()
10:         'row2array()
11:         comparador2013() '2013
            makeranking2013()
12:         novoprioridade2013()
            showresults()
            'MsgBox("no final, resultado1.ranking = " & resultado1.ranking)
            'MsgBox("no final, resultado2.ranking = " & resultado2.ranking)
            'MsgBox("no final, resultado3.ranking = " & resultado3.ranking)
            'MsgBox("no final, resultado4.ranking = " & resultado4.ranking)

13:         'AgruparAviados()
14:         'AvisarDespachos()
15:         'TresPresc()
16:         'limparzeros()
18:         'mostrarcruz()
19:         'verifgen()
1901:
1905:       limpo = False
1906:       'mostrardiab()
1907:
            comparado = True
            inicializado = False
            read_only()
20:     End If

SAIR:
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Comparar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Function calculo2(ByVal org As Short, ByVal gen As Boolean, ByVal comp2 As Double, ByVal intermedio2 As Double, ByVal pvp2 As Double, ByVal top5 As Double) As Double
1:      On Error GoTo MOSTRARERRO
        org = "1"
2:      If Not IsNothing(org) Then
3:          Select Case org
                Case 1 ', 2 'tipo 10
                    If comp2 > 0 Then
101:                    If Not IsNothing(intermedio2) Then
5:                          calculo2 = System.Math.Min(System.Math.Round(intermedio2 * comp2, 2), pvp2)
102:                    End If
                    Else
                        calculo2 = 0
                    End If
6:              Case 46 'tipo 17
7:                  If Not IsNothing(intermedio2) Then
103:                    calculo2 = System.Math.Min(System.Math.Round(intermedio2 * comp2, 2), pvp2)
104:                End If
8:              Case 42 'tipo 12
9:                  If Not IsNothing(intermedio2) Then
105:                    calculo2 = System.Math.Round(intermedio2, 2)
106:                End If
10:             Case 41 'tipo 11
11:                 If comp2 > 0 Then
12:                     If Not IsNothing(intermedio2) Then
107:                        calculo2 = System.Math.Round(intermedio2, 2)
108:                    End If
13:                 Else
14:                     calculo2 = 0
15:                 End If
16:             Case 67 'tipo 13
17:                 If comp2 > 0 Then
18:                     If Not IsNothing(intermedio2) Then
109:                        calculo2 = System.Math.Round(intermedio2, 2)
110:                    End If
19:                 Else
20:                     calculo2 = 0
21:                 End If
22:             Case 23, 24, 25 'diabetes - não sei se 24 e 25 também são assim mas já fica
23:                 If Not IsNothing(intermedio2) Then
111:                    calculo2 = System.Math.Round(intermedio2 * comp2, 2)
112:                End If
24:             Case 48 ', 57 'tipo 15
25:                 If comp2 > 0 Then
                        If pvp2 <= top5 And top5 > 0 Then
                            If Not IsNothing(intermedio2) Then
                                calculo2 = System.Math.Min(System.Math.Round(tectocomp * intermedio2, 2), pvp2)
                            End If
                        Else
115:                        If Not IsNothing(intermedio2) Then
29:                             calculo2 = System.Math.Min(System.Math.Round((System.Math.Min(tectocomp, (comp2 + 0.15))) * intermedio2, 2), pvp2)
116:                        End If
30:                     End If
31:                 Else
32:                     calculo2 = 0
33:                 End If
34:             Case 45 ', 59
117:                If Not IsNothing(intermedio2) Then
35:                     calculo2 = System.Math.Min(System.Math.Round(intermedio2 * (System.Math.Max(comp2, portcomp2)), 2), pvp2)
118:                End If
36:             Case 49 ', 68
37:                 If comp2 > 0 Then
                        If pvp2 <= top5 And top5 > 0 Then
                            If Not IsNothing(intermedio2) Then
                                calculo2 = System.Math.Min(System.Math.Round((System.Math.Min(System.Math.Max(tectocomp, (portcomp2 + 0.15)), tectocomp)) * intermedio2, 2), pvp2)
                            End If
                        Else
                            If Not IsNothing(intermedio2) Then
                                calculo2 = System.Math.Min(System.Math.Round((System.Math.Min(tectocomp, System.Math.Max((portcomp2 + 0.15), (comp2 + 0.15))) * intermedio2), 2), pvp2)
                            End If
                            : End If
                    Else
                        calculo2 = 0
                    End If
42:                 'Case 12 'SAMS
123:                '   If Not IsNothing(intermedio2) Then
43:                 '                 calculo2 = intermedio2 * (0.9)
124:                '                End If
44:                 End Select
45:
46:         'não tenho nada para o tipo 19(47) nem para os organismos
            '13_CGD, 
            '25_SAMSq, 
47:         '09[=(75% - 01)], 
            'CA, 
            'r1, r3, 
            '85, 87, h1[=(90% - 02)], 
            'j1, j7
            'o1, 
            'aa, ab, 
            'xv, 
            'fm, 
            '19
            'ds
            'SF[=(100% - 01)], SG[=(100% - 45)], SH[=(100% - 48)], SI[=(100% - 49)], 
48:     End If
49:     Exit Function
MOSTRARERRO:
        MsgBox("SUB calculo2: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Function SomarPVP2()
1:      On Error GoTo MOSTRARERRO
        If naoindicar = False Then
            'pvp11v é string
            'pvp11val 'e double
2:          Dim somaPVP2 As Double
3:          If pvp11.Text <> "" Then
4:              pvp11v = Replace(pvp11.Text.ToString, ".", ",")
5:              pvp11val = Convert.ToDouble(pvp11v)
6:          Else : pvp11val = 0
7:          End If
8:          If pvp22.Text <> "" Then
9:              pvp22v = Replace(pvp22.Text.ToString, ".", ",")
10:             pvp22val = Convert.ToDouble(pvp22v)
11:         Else : pvp22val = 0
12:         End If
13:         If pvp33.Text <> "" Then
14:             pvp33v = Replace(pvp33.Text.ToString, ".", ",")
15:             pvp33val = Convert.ToDouble(pvp33v)
16:         Else : pvp33val = 0
17:         End If
18:         If pvp44.Text <> "" Then
19:             pvp44v = Replace(pvp44.Text.ToString, ".", ",")
20:             pvp44val = Convert.ToDouble(pvp44v)
21:         Else : pvp44val = 0
22:         End If
23:         somaPVP2 = pvp11val + pvp22val + pvp33val + pvp44val
34:         SomarPVP2 = somaPVP2
        Else
            SomarPVP2 = 0
        End If
35:     Exit Function

MOSTRARERRO:
        MsgBox("SUB somarpvp2: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Function SomarComp2()
1:      On Error GoTo MOSTRARERRO
        If naoindicar = False Then
2:          Dim somaComp2 As Double
3:          If comp11.Text <> "" Then
4:              comp11v = Replace(comp11.Text, ".", ",")
5:              comp11val = Convert.ToDouble(comp11v)
6:          Else : comp11val = 0
7:          End If
8:          If comp22.Text <> "" Then
9:              comp22v = Replace(comp22.Text, ".", ",")
10:             comp22val = Convert.ToDouble(comp22v)
11:         Else : comp22val = 0
12:         End If
13:         If comp33.Text <> "" Then
14:             comp33v = Replace(comp33.Text, ".", ",")
15:             comp33val = Convert.ToDouble(comp33v)
16:         Else : comp33val = 0
17:         End If
18:         If comp44.Text <> "" Then
19:             comp44v = Replace(comp44.Text, ".", ",")
20:             comp44val = Convert.ToDouble(comp44v)
21:         Else : comp44val = 0
22:         End If
33:         somaComp2 = comp11val + comp22val + comp33val + comp44val
34:         SomarComp2 = somaComp2
        Else
            SomarComp2 = 0
        End If
35:     Exit Function

MOSTRARERRO:
        MsgBox("SUB somarcomp2: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function





    Sub indicar2(ByVal which As Short)
1:      On Error GoTo MOSTRARERRO
        If naoindicar = False Then
2:          If Not IsNothing(codigorow) Then
3:              Select Case which
                    Case 1
4:                      If av1.mostrado = "true" Then
5:
14:                         compdois = (a1row(6) * 0.01)
15:                         portcomp11 = portcompdois
16:                         intermedio2 = a1row(25)
17:                         pvp11.Text = intermedio2
18:                         pr11 = Replace(a1row(24), ".", ",")
19:                         If organismo = 48 Or organismo = 49 Then    'útil quando existe PRE
20:                             pr11 = taxapr * pr1                      ' possível fazer PRE=PR+20%, PRE=PR+25%, etc
21:                         End If
22:                         pr2 = pr11
23:                         If pr2 > 0 Then
                                intermedio2 = pr2    'existindo PR, o PVP não interessa pois o cálculo é sobre o PR directamente
24:                             'o de baixo era quando era usado PVP se PVP<PR
                                'intermedio2 = System.Math.Min(intermedio2, pr2)
25:                         End If
                            'tempcalc2 = a1row(25)  usei para limitar aqui a compSNS ao valor do PVP e fazia min(calculo2,tempcalc2)
26:                         comp11.Text = calculo2(organismo, gen, compdois, intermedio2, a1row(25), a1row(17)) '22 é pvpmenos1 e 19 é top5
                            'comp11.Text = System.Math.Min((System.Math.Round(calculo2(organismo, gen, compdois, intermedio2), 2)), tempcalc2)
                            'comp11.Text = System.Math.Round(calculo2(organismo, gen, compdois, intermedio2), 2)
27:                     End If
28:                 Case 2
29:                     If av2.mostrado = "true" Then
30:
39:                         compdois = (a2row(6) * 0.01)
40:                         portcompdois = portcompdois
41:                         intermedio2 = a2row(25)
42:                         pvp22.Text = intermedio2
43:                         pr22 = Replace(a2row(24), ".", ",")
44:                         If organismo = 48 Or organismo = 49 Then
45:                             pr22 = taxapr * pr2
46:                         End If
47:                         pr2 = pr22
48:                         If pr2 > 0 Then
                                intermedio2 = pr2
49:                             'o de baixo era quando era usado PVP se PVP<PR
                                'intermedio2 = System.Math.Min(intermedio2, pr2)
50:                         End If
                            'tempcalc2 = a2row(25)
                            comp22.Text = calculo2(organismo, gen, compdois, intermedio2, a2row(25), a2row(17)) '17 é pvp e 15 é top5
51:                         'compdois.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, compdois, intermedio2), 2)), tempcalc2)
                            'compdois.Text = System.Math.Round(calculo2(organismo, gen, compdois, intermedio2), 2)
52:                     End If
53:                 Case 3
54:                     If av3.mostrado = "true" Then
55:                         compdois = (a3row(6) * 0.01)
65:                         portcomp3 = portcompdois
66:                         intermedio2 = a3row(25)
67:                         pvp33.Text = intermedio2
68:                         pr33 = Replace(a3row(24), ".", ",")
69:                         If organismo = 48 Or organismo = 49 Then
70:                             pr33 = taxapr * pr3
71:                         End If
72:                         pr2 = pr33
73:                         If pr2 > 0 Then
                                intermedio2 = pr2
                                'o de baixo era quando era usado PVP se PVP<PR
74:                             'intermedio2 = System.Math.Min(intermedio2, pr2)
75:                         End If
                            'tempcalc2 = a3row(25)
                            comp33.Text = calculo2(organismo, gen, compdois, intermedio2, a3row(25), a3row(17)) '17 é pvp e 15 é top5
76:                         'comp3.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, compdois, intermedio2), 2)), tempcalc2)
                            'comp3.Text = System.Math.Round(calculo2(organismo, gen, compdois, intermedio2), 2)
77:                     End If
78:                 Case 4
79:                     If av4.mostrado = "true" Then
80:
89:                         compdois = (a4row(6) * 0.01)
90:                         portcomp4 = portcompdois
91:                         intermedio2 = a4row(25)
92:                         pvp44.Text = intermedio2
93:                         pr44 = Replace(a4row(24), ".", ",")
94:                         If organismo = 48 Or organismo = 49 Then
95:                             pr44 = taxapr * pr4
96:                         End If
97:                         pr2 = pr44
98:                         If pr2 > 0 Then
                                intermedio2 = pr2
                                'o de baixo era quando era usado PVP se PVP<PR
99:                             'intermedio2 = System.Math.Min(intermedio2, pr2)
100:                        End If
                            'tempcalc2 = a4row(25)
                            comp44.Text = calculo2(organismo, gen, compdois, intermedio2, a4row(25), a4row(17)) '17 é pvp e 15 é top5
101:                        'comp4.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, compdois, intermedio2), 2)), tempcalc2)
                            'comp4.Text = System.Math.Round(calculo2(organismo, gen, compdois, intermedio2), 2)
102:                    End If
103:
153:                    End Select
154:        End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("sub indicar2: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub acharduplicadosA()
        Select Case AA
            Case Is = 4
                If oa1.code = oa2.code Then
                    If oa1.code <> oa3.code And oa1.code <> oa4.code Then
                        oad12.code = oa1.code
                        oa1.duplicado = 2
                        oa2.duplicado = 1
                    ElseIf oa1.code = oa3.code Then
                        If oa1.code <> oa4.code Then
                            oat123.code = oa1.code
                        ElseIf oa1.code = oa4.code Then
                            oaq.code = oa1.code
                        End If
                    ElseIf oa1.code = oa4.code Then
                        oat124.code = oa1.code
                    End If
                End If
                If oa1.code = oa3.code Then
                    If oa1.code <> oa2.code And oa1.code <> oa4.code Then
                        oad13.code = oa1.code
                        oa1.duplicado = 3
                        oa3.duplicado = 1
                    End If
                End If
                If oa1.code = oa4.code Then
                    If oa1.code <> oa3.code And oa1.code <> oa2.code Then
                        oad14.code = oa1.code
                        oa1.duplicado = 4
                        oa4.duplicado = 1
                    End If
                End If
                If oa2.code = oa3.code Then
                    If oa1.code <> oa2.code And oa2.code <> oa4.code Then
                        oad23.code = oa2.code
                        oa2.duplicado = 3
                        oa3.duplicado = 2
                    ElseIf oa2.code = oa4.code Then
                        oat234.code = oa2.code
                    End If
                End If
                If oa2.code = oa4.code Then
                    If oa1.code <> oa2.code And oa2.code <> oa3.code Then
                        oad24.code = oa2.code
                        oa2.duplicado = 4
                        oa4.duplicado = 2
                    End If
                End If
                If oa4.code = oa3.code Then
                    If oa3.code <> oa1.code And oa2.code <> oa3.code Then
                        oad34.code = oa3.code
                        oa4.duplicado = 3
                        oa3.duplicado = 4
                    End If
                End If
        End Select
    End Sub

    Sub acharduplicadosP()
        Select Case PP
            Case Is = 4
                If op1.code = op2.code Then
                    If op1.code <> op3.code And op1.code <> op4.code Then
                        opd12.code = op1.code
                        op1.duplicado = 2
                        op2.duplicado = 1
                    ElseIf op1.code = op3.code Then
                        If op1.code <> op4.code Then
                            opt123.code = op1.code
                        ElseIf op1.code = op4.code Then
                            opq.code = op1.code
                        End If
                    ElseIf op1.code = op4.code Then
                        opt124.code = op1.code
                    End If
                End If
                If op1.code = op3.code Then
                    If op1.code <> op2.code And op1.code <> op4.code Then
                        opd13.code = op1.code
                        op1.duplicado = 3
                        op3.duplicado = 1
                    End If
                End If
                If op1.code = op4.code Then
                    If op1.code <> op3.code And op1.code <> op2.code Then
                        opd14.code = op1.code
                        op1.duplicado = 4
                        op4.duplicado = 1
                    End If
                End If
                If op2.code = op3.code Then
                    If op1.code <> op2.code And op2.code <> op4.code Then
                        opd23.code = op2.code
                        op2.duplicado = 3
                        op3.duplicado = 2
                    ElseIf op2.code = op4.code Then
                        opt234.code = op2.code
                    End If
                End If
                If op2.code = op4.code Then
                    If op1.code <> op2.code And op2.code <> op3.code Then
                        opd24.code = op2.code
                        op2.duplicado = 4
                        op4.duplicado = 2
                    End If
                End If
                If op4.code = op3.code Then
                    If op3.code <> op1.code And op2.code <> op3.code Then
                        opd34.code = op3.code
                        op4.duplicado = 3
                        op3.duplicado = 4
                    End If
                End If
        End Select
    End Sub








    Private Sub SóIsoladosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SóIsoladosToolStripMenuItem.Click
        On Error GoTo MOSTRARERRO
        If filtrosoisolados = True Then
            filtrosoisolados = False
            SóIsoladosToolStripMenuItem.Checked = False
        Else
            filtrosoisolados = True
            SóIsoladosToolStripMenuItem.Checked = True
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub SóIsoladosToolStripMenuItem_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub nDCI(ByVal prescrito As Single, ByVal aviado As Single, ByVal h As Single)
        On Error GoTo MOSTRARERRO
        Dim PPP(3) As Object
        Dim AAA(3) As Object
        PPP(0) = op1
        PPP(1) = op2
        PPP(2) = op3
        PPP(3) = op4
        AAA(0) = oa1
        AAA(1) = oa2
        AAA(2) = oa3
        AAA(3) = oa4
        If AAA(aviado - 1).gh = 0 And PPP(prescrito - 1).gh = 0 And PPP(prescrito - 1).CNPEM = False Then 'sem obrigatoriedade de prescrição por dci 'sem obrigatoriedade de prescrição por dci
            marcadif(aviado + 6, h)
        Else
            marcadif(aviado, h)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("nDCI: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



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
        Dim ARRAYdciP(3) As String
        Dim iteradorMR As Single = 11

21:     Do While iterAA <= A
            iterPP = 1
22:         Do While iterPP <= P
                iteradorMR = Convert.ToSingle((iterPP) & (iterAA))
23:             'ranking do code c*********
24:             If arraycruzamentos(iteradorMR).code = True Then 'se mesmo código
25:                 arraycruzamentos(iteradorMR).ranking = "411123"

26:                 GoTo OK
27:             ElseIf arrayp(iterPP - 1).porCNPEM = False Then 'se não prescrito por cnpem
                    'será como se xcnpem =1 ou 3
272:                If arrayp(iterPP - 1).cnpem = 0 Then 'prescrito por code diferente e sem CNPEM
273:                    arraycruzamentos(iteradorMR).ranking = "2" 'comparar à antiga, pode estar certo ou errado
274:                Else 'prescrito por code diferente com CNPEM
275:                    If arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por code dif com mesmo cnpem (pode estar certo ou S)
276:                        'arraycruzamentos(iteradorMR).ranking = "2111" & arraycruzamentos(iteradorMR).marcadifsemhavergens & "3"
                            'GoTo Ok
                            '[se deixasse como está em cima, no caso da marcadif ser 1, o showresult teria de ser alterado para detectar como erro]
                            If arraycruzamentos(iteradorMR).marcadifsemhavergens = 2 Then
277:                            arraycruzamentos(iteradorMR).ranking = "211123"
                                GoTo OK
279:                        Else
280:                            arraycruzamentos(iteradorMR).ranking = "211103" 'marcadif (0 ou 1)
                                GoTo OK
281:                        End If
282:                    Else
28:                         arraycruzamentos(iteradorMR).ranking = "1"  'por code dif com cnpem dif - está errado -  vai seguir o restante
29:                     End If
291:                End If
292:            ElseIf arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por cnpem e igual 'está diferente no não lite
30:                 arraycruzamentos(iteradorMR).ranking = "311123"
31:                 GoTo OK
32:             Else 'prescrito por cnpem e diferente (xcnpem = 2 ou 0)
33:                 arraycruzamentos(iteradorMR).ranking = "0" 'prescrito por cnpem e diferente

34:             End If
35:
36:             'ranking do dci *d********
37:             If arraycruzamentos(iteradorMR).dci = True Then
38:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
39:             Else
40:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
41:             End If
42:
43:             'ranking da forma **f*******
44:             If arraycruzamentos(iteradorMR).forma = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'era diferente no não lite

45:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
46:             Else
47:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
48:             End If
49:
50:             'ranking da dose ***d******
51:             If arraycruzamentos(iteradorMR).dose = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'era diferente no não lite
52:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
53:             Else
54:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
55:             End If
56:
57:             'ranking da marca ****n*****
                arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).marcadifsemhavergens

63:
64:             'ranking da qty *****q****
                arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).xqty
65:
66:
OK:             'ranking da excexp a) ******a***
68:             'ranking da excexp b) *******b**
69:             'ranking da excexp c) ********c*
                If arrayp(iterPP - 1).porCNPEM = False Then 'só há excepções se não for prescrição por CNPEM
70:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).excepa & arraycruzamentos(iteradorMR).excepb & arraycruzamentos(iteradorMR).excepc
71:             Else 'prescrito por CNPEM
                    arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "555"
                End If

72:             'ranking do top5 *********t
73:             arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).top5
74:
75:
76:             iterPP = iterPP + 1
77:         Loop
78:         iterAA = iterAA + 1
79:     Loop

80:     Exit Sub
MOSTRARERRO:
        MsgBox("sub makeranking2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub comparador2013() '2013
1:      On Error GoTo MOSTRARERRO
        Dim posicaop1(3) As Single
        Dim posicaop2(3) As Single
        Dim posicaoa1(3) As Single
        Dim posicaoa2(3) As Single
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
        Dim iterador As Single = 11
33:     Do While iterA <= A
            iterP = 1
34:         Do While iterP <= P
                iterador = Convert.ToSingle((iterP) & (iterA))
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
827:                    End If
828:                    'compara antesp(iterP - 1) com antesa(iterA - 1)
829:                    If antesp(iterP - 1) <> antesa(iterA - 1) Then 'se a quantidade não numerica não é igual vai ver se é menor, maior e maior que 50% (independentemente CNPEM)
830:                        If Convert.ToInt16(antesp(iterP - 1)) > Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado < prescrito
831:                            arraycruz(iterador).xqty = 2
832:                        ElseIf Convert.ToInt16(antesp(iterP - 1)) > 1.5 * Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado > 150% prescrito
833:                            arraycruz(iterador).xqty = 0
834:                        ElseIf Convert.ToInt16(antesp(iterP - 1)) < 1.5 * Convert.ToInt16(antesa(iterA - 1)) And Convert.ToInt16(antesa(iterA - 1)) > Convert.ToInt16(antesp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
835:                            arraycruz(iterador).xqty = 1
                            Else
                                arraycruz(iterador).xqty = 3
8136:                       End If
8137:                   Else 'se antes é igual compara depoisp(iterP - 1) com depoisa(iterA - 1)
8138:                       If Convert.ToInt16(depoisp(iterP - 1)) > Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado < prescrito
8139:                           arraycruz(iterador).xqty = 2
8140:                       ElseIf Convert.ToInt16(depoisp(iterP - 1)) > 1.5 * Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado > 150% prescrito
8141:                           arraycruz(iterador).xqty = 0
8142:                       ElseIf Convert.ToInt16(depoisp(iterP - 1)) < 1.5 * Convert.ToInt16(depoisa(iterA - 1)) And Convert.ToInt16(depoisa(iterA - 1)) > Convert.ToInt16(depoisp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
8143:                           arraycruz(iterador).xqty = 1
                            Else
                                arraycruz(iterador).xqty = 3
8144:                       End If
8145:                   End If
8146:               End If
8148:           End If



8231:           If arrayA(iterA - 1).comp = arrayP(iterP - 1).comp Then
8232:               arraycruz(iterador).comp = True
8233:           Else
8234:               arraycruz(iterador).comp = False
8235:           End If

836:            If arrayA(iterA - 1).GH = arrayP(iterP - 1).GH Then
837:                arraycruz(iterador).GH = True
838:            Else
839:                arraycruz(iterador).GH = False
840:            End If

841:            If arrayA(iterA - 1).pr = arrayP(iterP - 1).pr Then
842:                arraycruz(iterador).pr = True
843:            Else
844:                arraycruz(iterador).pr = False
845:            End If

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

866:            If arrayA(iterA - 1).d4250 = arrayP(iterP - 1).d4250 Then
867:                arraycruz(iterador).d4250 = True
868:            Else
869:                arraycruz(iterador).d4250 = False
870:            End If

871:            If arrayA(iterA - 1).d1234 = arrayP(iterP - 1).d1234 Then
872:                arraycruz(iterador).d1234 = True
873:            Else
874:                arraycruz(iterador).d1234 = False
875:            End If

876:            If arrayA(iterA - 1).d21094 = arrayP(iterP - 1).d21094 Then
877:                arraycruz(iterador).d21094 = True
878:            Else
879:                arraycruz(iterador).d21094 = False
880:            End If

881:            If arrayA(iterA - 1).d10279 = arrayP(iterP - 1).d10279 Then
882:                arraycruz(iterador).d10279 = True
883:            Else
884:                arraycruz(iterador).d10279 = False
885:            End If

886:            If arrayA(iterA - 1).d10280 = arrayP(iterP - 1).d10280 Then
887:                arraycruz(iterador).d10280 = True
888:            Else
889:                arraycruz(iterador).d10280 = False
890:            End If

891:            If arrayA(iterA - 1).d10910 = arrayP(iterP - 1).d10910 Then
892:                arraycruz(iterador).d10910 = True
893:            Else
894:                arraycruz(iterador).d10910 = False
895:            End If

896:            If arrayA(iterA - 1).d14123 = arrayP(iterP - 1).d14123 Then
897:                arraycruz(iterador).d14123 = True
898:            Else
899:                arraycruz(iterador).d14123 = False
900:            End If

901:            If arrayA(iterA - 1).lei6 = arrayP(iterP - 1).lei6 Then
902:                arraycruz(iterador).lei6 = True
903:            Else
904:                arraycruz(iterador).lei6 = False
905:            End If

906:            If arrayA(iterA - 1).pvpmenos1 = arrayP(iterP - 1).pvpmenos1 Then
907:                arraycruz(iterador).pvpmenos1 = True
908:            Else
909:                arraycruz(iterador).pvpmenos1 = False
910:            End If

911:            If arrayA(iterA - 1).pvpmenos2 = arrayP(iterP - 1).pvpmenos2 Then
912:                arraycruz(iterador).pvpmenos2 = True
913:            Else
914:                arraycruz(iterador).pvpmenos2 = False
915:            End If

916:            If arrayA(iterA - 1).trocamarca = arrayP(iterP - 1).trocamarca Then
917:                arraycruz(iterador).trocamarca = True
918:            Else
919:                arraycruz(iterador).trocamarca = False
920:            End If
921:
922:            If arrayA(iterA - 1).cnpem = arrayP(iterP - 1).cnpem Then
923:                arraycruz(iterador).cnpem = True
924:            Else
925:                arraycruz(iterador).cnpem = False
926:            End If

927:            If arrayA(iterA - 1).nDCI = arrayP(iterP - 1).nDCI Then
928:                arraycruz(iterador).nDCI = True
929:            Else
930:                arraycruz(iterador).nDCI = False
931:            End If

932:            If arrayA(iterA - 1).pvp = arrayP(iterP - 1).pvp Then
933:                arraycruz(iterador).pvp = True
934:            Else
935:                arraycruz(iterador).pvp = False
936:            End If

84:             If arraycruz(iterador).pvp = False Then 'se o pvp não é igual vai ver se é menor, maior    [sem ser unitário]
85:                 If arrayP(iterP - 1).pvp > arrayA(iterA - 1).pvp Then 'se pvp aviado < prescrito
86:                     arraycruz(iterador).xpvp = 1
87:                 ElseIf arrayP(iterP - 1).pvp > arrayA(iterA - 1).pvp Then 'se pvp aviado > prescrito
88:                     arraycruz(iterador).xpvp = 2
89:                 End If
90:             Else
91:                 arraycruz(iterador).xpvp = 0 'pvp =   
92:             End If
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
132:
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
179:                If arraycruz(iterador).code = False Then 'And arraycruz(iterador).lab = False Then 'se prescrito e aviado tÊm nome, lab e código diferentes
                        arraycruz(iterador).excepb = 0
                    Else
                        arraycruz(iterador).excepb = 2
                    End If
                End If

                'código para dar valores aos excep c
                'mesmo código =3
                'mais barato=2
                'mesmo preço=1
                'mais caro=0

                If arraycruz(iterador).code = False Then 'And arraycruz(iterador).lab = False Then  'se prescrito e aviado têm codigo diferente
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


                'código para dar valores a marca diferente
                'com genéricos mesmo gen/marca = 2
                'sem genéricos mesma marca =1
                'sem genérico marca dif =0

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


                'código para avisar top5
                'se não tem top5 =3
                'se tem top5 e é mais barato =2
                'se tem top5 e é mesmo preço =1
                'se tem top5 e é mais caro =0
                If arrayP(iterP - 1).top5 > 0 Then 'prescrito tem top5
                    If arrayA(iterA - 1).top5 > 0 Then 'aviado tem top5
                        If arrayA(iterA - 1).pvp > arrayP(iterP - 1).top5 Then
                            arraycruz(iterador).top5 = 0
                        ElseIf arrayA(iterA - 1).pvp = arrayP(iterP - 1).top5 Then
                            arraycruz(iterador).top5 = 1
                        Else
                            arraycruz(iterador).top5 = 2
                        End If
                    End If
                Else
                    arraycruz(iterador).top5 = 3
                End If


169:            'End If
170:
180:
181:            iterP = iterP + 1
182:        Loop
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

        Dim arraycombo(3) As Object
        arraycombo(0) = ComboBox1
        arraycombo(1) = ComboBox2
        arraycombo(2) = ComboBox3
        arraycombo(3) = ComboBox4

        Dim arrayvalorcombopvp(3) As Object
        arrayvalorcombopvp(0) = valorcombopvp1
        arrayvalorcombopvp(1) = valorcombopvp2
        arrayvalorcombopvp(2) = valorcombopvp3
        arrayvalorcombopvp(3) = valorcombopvp4

632:    For letraA As Integer = 0 To A - 1
633:        arrayarow(letraA) = DS.infarmed.FindBycode(arrayaviado(letraA).codigo)
634:        arrayaarray(letraA).Add(arrayarow(letraA))
635:        If Not IsNothing(arrayarow(letraA)) Then
636:            arrayoa(letraA).code = arrayarow(letraA)(0)
637:            arrayoa(letraA).nome = arrayarow(letraA)(2)
638:            arrayoa(letraA).d10910 = arrayarow(letraA)(15)
639:            arrayoa(letraA).pvpmenos2 = arrayarow(letraA)(21)
640:            arrayoa(letraA).cnpem = arrayarow(letraA)(22)
641:            arrayoa(letraA).dci = LCase(arrayarow(letraA)(1))
70:             arrayoa(letraA).forma = arrayarow(letraA)(3)
71:             arrayoa(letraA).dose = arrayarow(letraA)(4)
72:             arrayoa(letraA).qty = Replace(arrayarow(letraA)(5).ToString, ".", ",")
73:             arrayoa(letraA).comp = arrayarow(letraA)(6)
74:             arrayoa(letraA).gh = arrayarow(letraA)(7)
75:             arrayoa(letraA).gen = arrayarow(letraA)(8)
76:             arrayoa(letraA).lab = arrayarow(letraA)(9)
77:             arrayoa(letraA).pvp = arrayarow(letraA)(23)
78:             arrayoa(letraA).pr = arrayarow(letraA)(24)
79:             arrayoa(letraA).d4250 = arrayarow(letraA)(10)
80:             arrayoa(letraA).d1234 = arrayarow(letraA)(11)
81:             arrayoa(letraA).d10279 = arrayarow(letraA)(13)
82:             arrayoa(letraA).d10280 = arrayarow(letraA)(14)
83:             arrayoa(letraA).d21094 = arrayarow(letraA)(12)
84:             arrayoa(letraA).lei6 = arrayarow(letraA)(19)
85:             arrayoa(letraA).d14123 = arrayarow(letraA)(16)
86:             arrayoa(letraA).dci_obr = arrayarow(letraA)(18)
87:             arrayoa(letraA).trocamarca = arrayarow(letraA)(20)
88:             arrayoa(letraA).top5 = arrayarow(letraA)(17)
89:             arrayoa(letraA).pvpmenos1 = arrayarow(letraA)(25)
                arrayoa(letraA).pvpmenos3 = arrayarow(letraA)(26)
                arrayoa(letraA).pvpmenos4 = arrayarow(letraA)(27)
                arrayoa(letraA).pvpmenos5 = arrayarow(letraA)(28)

                Select Case mesactual
                    Case Is = 1, 4, 7, 10
                        'usam-se 4 (o actual e os 3 anteriores)
                        ReDim arrayoa(letraA).arraypvp(3)
                        arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3}
                    Case Is = 2, 5, 8, 11
                        'usam-se 5 (o actual e os 4 anteriores)
                        ReDim arrayoa(letraA).arraypvp(4)
                        arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3, arrayoa(letraA).pvpmenos4}

                    Case Is = 3, 6, 9, 12
                        'usam-se 6 (o actual e os 5 anteriores)
                        ReDim arrayoa(letraA).arraypvp(5)
                        arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3, arrayoa(letraA).pvpmenos4, arrayoa(letraA).pvpmenos5}

                End Select
                arraycombo(letraA).DataSource = arrayoa(letraA).arraypvp
                arrayvalorcombopvp(letraA) = True
                valorcombopvp1 = arrayvalorcombopvp(0) 'senão não  funciona não sei porquê
                valorcombopvp2 = arrayvalorcombopvp(1) 'senão não  funciona não sei porquê
                valorcombopvp3 = arrayvalorcombopvp(2) 'senão não  funciona não sei porquê
                valorcombopvp4 = arrayvalorcombopvp(3) 'senão não  funciona não sei porquê
                ComboBox1.DataSource = oa1.arraypvp
                ComboBox2.DataSource = oa2.arraypvp
                ComboBox3.DataSource = oa3.arraypvp
                ComboBox4.DataSource = oa4.arraypvp

                If IsNumeric(arrayoa(letraA).qty) Then
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
910:            arrayop(letraP).d10910 = arrayprow(letraP)(15)
911:            arrayop(letraP).pvpmenos2 = arrayprow(letraP)(21)
912:            arrayop(letraP).cnpem = arrayprow(letraP)(22)
913:            arrayop(letraP).dci = LCase(arrayprow(letraP)(1))
170:            arrayop(letraP).forma = arrayprow(letraP)(3)
171:            arrayop(letraP).dose = arrayprow(letraP)(4)
172:            arrayop(letraP).qty = Replace(arrayprow(letraP)(5).ToString, ".", ",")
173:            arrayop(letraP).comp = arrayprow(letraP)(6)
174:            arrayop(letraP).gh = arrayprow(letraP)(7)
175:            arrayop(letraP).gen = arrayprow(letraP)(8)
176:            arrayop(letraP).lab = arrayprow(letraP)(9)
177:            arrayop(letraP).pvp = arrayprow(letraP)(23)
178:            arrayop(letraP).pr = arrayprow(letraP)(24)
179:            arrayop(letraP).d4250 = arrayprow(letraP)(10)
180:            arrayop(letraP).d1234 = arrayprow(letraP)(11)
181:            arrayop(letraP).d10279 = arrayprow(letraP)(13)
182:            arrayop(letraP).d10280 = arrayprow(letraP)(14)
183:            arrayop(letraP).d21094 = arrayprow(letraP)(12)
184:            arrayop(letraP).lei6 = arrayprow(letraP)(19)
185:            arrayop(letraP).d14123 = arrayprow(letraP)(16)
186:            arrayop(letraP).dci_obr = arrayprow(letraP)(18)
187:            arrayop(letraP).trocamarca = arrayprow(letraP)(20)
188:            arrayop(letraP).top5 = arrayprow(letraP)(17)
189:            arrayop(letraP).pvpmenos1 = arrayprow(letraP)(25)
                arrayop(letraP).pvpmenos3 = arrayprow(letraP)(26)
                arrayop(letraP).pvpmenos4 = arrayprow(letraP)(27)
                arrayop(letraP).pvpmenos5 = arrayprow(letraP)(28)
                arrayop(letraP).pvpexcepC = PVPmax(arrayop(letraP).pvp, arrayop(letraP).pvpmenos1, arrayop(letraP).pvpmenos2, arrayop(letraP).pvpmenos3, arrayop(letraP).pvpmenos4, arrayop(letraP).pvpmenos5)
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
234:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub irbuscar2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
52:     Dim iteradorQual1 As SByte = 11
53:     Dim iteradorQual2 As SByte = 11
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
128:                cruzamentocerto(iterAAA - 1).ranking = "0000000000" 'AVIADO NÃO PRESCRiTO
129:                'Case Is = 1 'só um prescrito, ligação directa 'pus plickas para o caso de haver mais aviados e assim vai à regra geral
130:                '   cruzamentocerto(iterAAA - 1).qualP = iterPPP
131:                '  cruzamentocerto(iterAAA - 1).ranking = arraycroces(iteradorPrior).ranking
132:                ' iteradorQual1 = Convert.ToSByte(cruzamentocerto(iterAAA - 1).qualP & iterAAA)
133:                'anular(iteradorQual1, iteradorQual2, iterAAA, arraycroces, cruzamentocerto)

134:            Case Else
135:                SUBROTINA2(iterPPP, iterAAA, iteradorPrior, iteradorQual1, iteradorQual2, arraycroces, cruzamentocerto, quantosdciAnosP(iterAAA - 1), quantosdciAnosA(iterAAA - 1))
136:                End Select
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

151:
152:        'MsgBox(iterAAA - 1 & " " & cruzamentocerto(iterAAA - 1).ranking)
159:        iterAAA = iterAAA + 1
160:    Loop 'final do while iterAAA

161:    'para a CC podem ser aviados mais que os prescritos
162:    'COLOCAR a verificação de mesmo dci/forma/dose entre vários aviados (chamada da sub)
        'ou aqui ou dentro do loop ou no fim da subrotina (<= deve ter melhor acesso a todas as variaveis)
        'se mesmo dci/forma/dose e há menos prescritos nessas condições e desnivel de 1,5x entre total qty prescrita e cada um desses aviados
        'então soma quantidade aviada desses prescritos e compara com a quantidade total prescrita.
        'se <= 1,5 => OK para 3808 e necessita justificação para mim (e)
        'else erro (h)
163:    Exit Sub
MOSTRARERRO:
        MsgBox("sub novoprioridade2013: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'acrescentar um verificador que por default é verde
    'fica amarelo se um dos results é amarelo
    'fica vermelho
    'fica laranja
    Sub showresults()
        On Error GoTo MOSTRARERRO
        Dim show(3) As embalagem.cruzamento  'assim além do ranking posso ir buscar o qualp e outras coisas
3:      show(0) = resultado1
4:      show(1) = resultado2
5:      show(2) = resultado3
6:      show(3) = resultado4
7:      Dim reresult(3) As Object
8:      reresult(0) = result1
9:      reresult(1) = result2
10:     reresult(2) = result3
11:     reresult(3) = result4
        Dim doaviado(3) As Object
        doaviado(0) = oa1
        doaviado(1) = oa2
        doaviado(2) = oa3
        doaviado(3) = oa4
        Dim doprescrito(3) As Object
        doprescrito(0) = op1
        doprescrito(1) = op2
        doprescrito(2) = op3
        doprescrito(3) = op4
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
        Dim arrayc(3) As Object
        arrayc(0) = C_PVP_1
        arrayc(1) = C_PVP_2
        arrayc(2) = C_PVP_3
        arrayc(3) = C_PVP_4
        Dim dciAnosP(3) As Object
        dciAnosP(0) = quantosDCIa1nosP
        dciAnosP(1) = quantosDCIa2nosP
        dciAnosP(2) = quantosDCIa3nosP
        dciAnosP(3) = quantosDCIa4nosP

        Dim linha As Integer = 0
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
31:                             Mid(show(toshow).ranking, secnpem + 1, 1) = "1"
32:                         End If
                        End If

                        'meter o xqty em vez do qty caso (5) <>3
33:                 Next
334:            End If
34:         End If
341:
342:        Dim tempqty As Single = 0
343:
344:        'MsgBox(show(toshow).ranking)
345:        With reresult(toshow) 'richtextbox
346:            Dim dizera As String = ""
347:            Dim dizerp As String = ""
348:            'MsgBox(doprescrito(show(toshow).qualP - 1).code)
                If show(toshow).qualP > 0 Then
349:                If doprescrito(show(toshow).qualP - 1).code > 50000000 Then 'prescrito por cnpem
350:                    dizerp = UCase(doprescrito(show(toshow).qualP - 1).dci.substring(0, 1)) & doprescrito(show(toshow).qualP - 1).dci.substring(1, (Len(doprescrito(show(toshow).qualP - 1).dci) - 1))
351:                Else
352:                    dizerp = UCase(doprescrito(show(toshow).qualP - 1).nome.substring(0, 1)) & doprescrito(show(toshow).qualP - 1).nome.substring(1, (Len(doprescrito(show(toshow).qualP - 1).nome) - 1))
353:                End If
354:            End If
355:            If show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(5, 1) = 2 Then
356:                tempqty = show(toshow).ranking.ToString.Substring(5, 1)
357:                Mid(show(toshow).ranking, 6, 1) = "0"
358:            ElseIf show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(5, 1) = 3 Then
359:                tempqty = show(toshow).ranking.ToString.Substring(5, 1)
3561:           End If

3562:           If show(toshow).ranking = "0000000000" Then
412:                If show(toshow).CNPEM > 0 Then
4121:                   For ixp As Integer = 0 To P - 1
4122:                       If doprescrito(ixp).cnpem = doaviado(toshow).CNPEM Then
4123:                           .text = doaviado(toshow).nome & " repetido"
4124:                           Exit For
4125:                       End If
4126:                       .text = "y" & toshow + 1 & " = -> " & doaviado(toshow).nome & " não prescrito"
4127:                   Next
4128:               Else
4129:                   For iixp As Integer = 0 To P - 1
4130:                       If doprescrito(iixp).dci = doaviado(toshow).dci Then
4131:                           .text = doaviado(toshow).nome & " repetido"
4132:                           Exit For
4133:                       End If
4134:                       .text = "y" & toshow + 1 & " = -> " & doaviado(toshow).nome & " não prescrito"
4135:                   Next
4136:               End If
4144:               show(toshow).erro = True
415:                verificador = Color.Red
                    .selectionstart = 0
                    .selectionlength = Len(reresult(toshow).text)
                    .selectionbackcolor = verificador
4151:           Else
416:                For posrank As Integer = 0 To 5


42:                     If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 AndAlso show(toshow).erro = False Then 'cnpem<>, dci<>, formaCNPEM<>, doseCNPEM<>, marca<>, qtyCNPEM<>
43:                         'MsgBox(show(toshow).ranking)
431:                        Select Case posrank  'ir concatenado += na richtextbox
                                Case Is = 0
433:                                'cnpem diferente mas como não diz o porquê não mostra nada
434:                            Case Is = 1 'dci <>
                                    .text += show(toshow).qualP & "y" & toshow + 1 & " = " & dizerp & " -> " & doaviado(toshow).dci & vbNewLine
                                    show(toshow).erro = True
437:                                verificador = Color.Red
438:                            Case Is = 2 'forma <>
439:                                .text += show(toshow).qualP & "f" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).FORMA & " -> " & doaviado(toshow).FORMA & vbNewLine
440:                                show(toshow).erro = True
441:                                verificador = Color.Red
442:                            Case Is = 3 'dose <>
443:                                .text += show(toshow).qualP & "g" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).dose & " -> " & doaviado(toshow).dose & vbNewLine
444:                                show(toshow).erro = True
445:                                verificador = Color.Red
446:                            Case Is = 4 'marca <>
447:                                If Not doprescrito(show(toshow).qualP - 1).code > 50000000 Then 'para não mostrar marca dif quando prescrito por cnpem
448:                                    .text += show(toshow).qualP & "s" & toshow + 1 & " = " & doprescrito(show(toshow).qualP - 1).nome & " -> " & doaviado(toshow).nome & vbNewLine
449:                                    show(toshow).erro = True
                                        verificador = Color.Red
                                    Else

450:                                End If
451:                                'show(toshow).erro = True 'como costuma estar mal mais vale não impedir de mostrar qtys
452:                            Case Is = 5
453:
454:                                Select Case tempqty 'ver se h) ou L
                                        Case Is = 0 'h)
456:                                        .text += show(toshow).qualP & "h" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " " & vbNewLine
457:                                        show(toshow).erro = True
                                            verificador = Color.Red
458:                                    Case Is = 1 'x<y<150%
459:                                        .text += show(toshow).qualP & "h" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " " & vbNewLine
460:                                        show(toshow).erro = True
                                            verificador = Color.Red
4601:                                   Case Is = 2 'L
4602:                                       .text += show(toshow).qualP & "L" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " " & vbNewLine
4603:                                       show(toshow).erro = True
                                            If Not verificador = Color.Red Then
                                                verificador = Color.Orange
                                            End If
4604:                                   Case Else
4605:                                       End Select
4606:                           Case Else
4607:                               End Select
4608:                       .selectionstart = 0
4609:                       .selectionlength = Len(reresult(toshow).text)
4610:                       .selectionbackcolor = verificador
44:                     Else

45:                     End If

46:                 Next
461:            End If


462:            If show(toshow).erro = False Then   'a seguir acrescentei 0 a 17/12/2013
                    If show(toshow).ranking.ToString.Substring(0, 1) = 1 Or show(toshow).ranking.ToString.Substring(0, 1) = 0 Then 'prescrito por code com CNPEM dif e não encontrou erros
                        verificador = Color.Red
                        show(toshow).erro = True
                        .text += dizerp & " " & doprescrito(show(toshow).qualP - 1).cnpem & " " & " CNPEM " & " " & doaviado(toshow).cnpem
464:                    .selectionstart = 0
465:                    .selectionlength = 26 + Len(dizerp)
                        .selectionbackcolor = verificador
                    Else
463:                    .text += " OK " & vbNewLine 'Environment.NewLine pode ser usado em vez do vbnewline
                        .selectionstart = 0
                        .selectionlength = 4
466:                    .selectionbackcolor = Color.Green
467:                    'verificador = Color.Green
                    End If
468:            End If
469:
470:
57:
58:
62:         End With


            'verificar de aviados mais de dois iguais - só acusa se támbém mais de dois iguais prescritos (para não acusar desdobramentos)
            If detectarmaisdedoisaviados() = True Then
                verificador = Color.Red
                If Not excepBox.Text.Contains("mais de 2 iguais") Then
                    excepBox.Text += "mais de 2 iguais" & vbNewLine
                End If
            End If



            With excepBox
                If show(toshow).ranking.ToString.Substring(8, 1) = 1 Or 2 Then 'mb ou mp  
471:                If Not (show(toshow).ranking.ToString.Substring(0, 1) = "0" Or "1") Then 'para não mostrar mais barato quando prescrito por cnpem
472:                    If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome And show(toshow).ranking.ToString.Substring(5, 1) < 3) Then 'para não acusar excep quando mesmo medicamento com L ou H
                            If verificador = Color.Green Then
                                verificador = Color.Yellow
                            End If

                            .Text += show(toshow).qualP & "t " & toshow + 1 & vbNewLine '& " "
473:                        ' .SelectionStart = Len(reresult(toshow).text.ToString)
474:                        ' .SelectionLength = 5
475:                        ' .SelectionBackColor = verificador
54:                     End If
55:                 End If
56:             End If


47:             ' For posrank As Integer = 6 To 9
48:             'If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 Then 'a), b), c) ou (>)
49:             '             If verificador = Color.Green Then
                'verificador = Color.Yellow
4901:           'End If
4902:
4903:           If show(toshow).erro = False Then
4904:               If show(toshow).ranking.ToString.Substring(6, 1) = 0 Or show(toshow).ranking.ToString.Substring(7, 1) = 0 Or show(toshow).ranking.ToString.Substring(8, 1) = 0 Then
4905:                   If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome And show(toshow).ranking.ToString.Substring(5, 1) < 3) Then 'para não acusar excep quando mesmo medicamento com L ou H
4906:                       'linha = linha + 1
4907:                       .Text += show(toshow).qualP & " " & doaviado(toshow).dci & " " & toshow + 1 & vbNewLine
4908:                       reresult(toshow).text += show(toshow).qualP & " c) " & toshow + 1 & vbNewLine
4909:                       reresult(toshow).selectionstart = reresult(toshow).Text.IndexOf(reresult(toshow).Lines(1))
4910:                       reresult(toshow).selectionlength = 6
4911:                       reresult(toshow).selectionbackcolor = Color.Yellow
4912:                       arrayc(toshow).Text = doprescrito(show(toshow).qualP - 1).PVPexcepc
4913:                       If verificador = Color.Green Then
4914:                           verificador = Color.Yellow
4915:                       End If
4916:                   End If
4917:               End If
4918:           End If
4919:           ' .Lines = New String() {linha - 1}
50:             ' .SelectionStart = 0
                ' .SelectionLength = Len(doaviado(toshow).dci) + 4
                'If posrank = 9 Then
                '.SelectionLength = 8
                'End If
                '.SelectionBackColor = Color.Yellow
51:             'End If
52:             ' Next
53:         End With

            With top5Box
                For posrank As Integer = 6 To 9
                    If show(toshow).erro = False Then
                        If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 Then 'a), b), c) ou (>)
                            'os das excepções estão na excepbox
                            Select Case posrank
                                Case Is = 9
                                    .Text += ">top5 " & toshow + 1 & vbNewLine '& " "
                            End Select
                            .SelectionStart = 0
                            .SelectionLength = 7
                            .SelectionBackColor = verificador
                        End If
                    End If
                Next
            End With
        Next

        Select Case verificador
            Case Is = Color.Green
                PictureBox2.Load("green.png")
            Case Is = Color.Yellow
                PictureBox2.Load("yellow.png")
            Case Is = Color.Orange
                PictureBox2.Load("orange.png")
            Case Is = Color.Red
                PictureBox2.Load("red.png")
        End Select

        PictureBox1.Hide()
        PictureBox2.Show()
        excepBox.BringToFront()
        'falta ARRANJAR SíTIO PARA ISTO
146:    'If aceitarduplicados = True Then
150:    '   'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
165:    '  For AAA As Integer = 1 To A
169:    '    If AAA <> iterAAA Then
        'If (arrayoa(AAA - 1).cnpem > 0 And (arrayoa(AAA - 1).cnpem = arrayoa(iterAAA - 1).cnpem)) Or (arrayoa(AAA - 1).code = arrayoa(iterAAA - 1).code) Then
170:    '    If iterAAA > AAA Then 'para não fazer duas vezes e desfazer o que estava feito
        'cruzamentocerto(iterAAA - 1).ranking = cruzamentocerto(AAA - 1).ranking
        'End If
172:    'End If
173:    'End If
174:    'Next
        'End 
        ' For Each C As Control In Controls
        ' If TypeOf C Is Label Then
        ' DirectCast(C, Label).BackColor = verificador
        '  End If
        '  If TypeOf C Is RichTextBox Then
        ' DirectCast(C, RichTextBox).BackColor = verificador
        ' End If
        ' If TypeOf C Is GroupBox Then
        ' DirectCast(C, GroupBox).BackColor = verificador
        ' End If
        '   If C.HasChildren Then
        ' UpdateLabelFG(C.Controls, verificador)
        ' End If
        ' Next
        ' hora.BackColor = SystemColors.GradientActiveCaption
        ' Me.BackColor = verificador
        resultadomostrado = True
        Exit Sub
MOSTRARERRO:
        MsgBox("sub showresults: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    ' Private Sub UpdateLabelFG(ByVal controls As ControlCollection, ByVal fgColor As Color)
    '     If controls Is Nothing Then Return
    '     For Each C As Control In controls
    '         If TypeOf C Is Label Then DirectCast(C, Label).BackColor = verificador
    '         If C.HasChildren Then UpdateLabelFG(C.Controls, fgColor)
    '     Next
    ' End Sub



    Sub SUBROTINA(ByVal ITERppp As Single, ByVal ITERaaa As Single, ByVal ITERadorprior As Single, ByVal iteradorQual1 As SByte, ByVal iteradorQual2 As SByte, ByVal arraycroces As Array, ByVal cruzamentocerto As Array, ByVal quantosdciAnosP As Single, ByVal quantosdciAnosA As Single)
1:      On Error GoTo MOSTRARERRO
        cruzamentocerto(0) = resultado1
        cruzamentocerto(1) = resultado2
        cruzamentocerto(2) = resultado3
        cruzamentocerto(3) = resultado4
2:      Dim ordenados(3) As Array
3:      ordenados(0) = ordenado1
4:      ordenados(1) = ordenado2
5:      ordenados(2) = ordenado3
6:      ordenados(3) = ordenado4
7:      '   quantosdciAnosP(0) = quantosDCIa1nosP
84:     '  quantosdciAnosP(1) = quantosDCIa2nosP
85:     ' quantosdciAnosP(2) = quantosDCIa3nosP
86:     'quantosdciAnosP(3) = quantosDCIa4nosP

88:     'quantosdciAnosA(0) = quantosDCIa1
89:     'quantosdciAnosA(1) = quantosDCIa2
90:     'quantosdciAnosA(2) = quantosDCIa3
91:     'quantosdciAnosA(3) = quantosDCIa4

92:     'If arraycroces(ITERadorprior).anulado = False Then



93:     ITERppp = 1
94:     Do While ITERppp <= P
361:
363:        ITERadorprior = Convert.ToSingle((ITERppp) & (ITERaaa))
37:         If arraycroces(ITERadorprior).anulado = False Then
38:             If arraycroces(ITERadorprior).ranking.ToString.Substring(1) >= cruzamentocerto(ITERaaa - 1).ranking.ToString.Substring(1) Then
39:                 If Not (quantosdciAnosP - quantosdciAnosA) < 0 Then
390:                    cruzamentocerto(ITERaaa - 1).qualP = ITERppp
391:                    cruzamentocerto(ITERaaa - 1).ranking = arraycroces(ITERadorprior).ranking
393:                    iteradorQual1 = Convert.ToSByte(cruzamentocerto(ITERaaa - 1).qualP & ITERaaa)
394:                    iteradorQual2 = iteradorQual1
395:                Else         'mais aviados que prescritos 
396:                    'ordeno por ranking

397:                    For o As Integer = 0 To quantosdciAnosA - 1
398:                        concatenado = cruzamentocerto(ITERaaa - 1).ranking & ITERadorprior
399:                        ordenados(ITERaaa - 1)(o) = concatenado
4001:                       '  MsgBox((ITERaaa - 1) & " " & o & " " & ordenados(ITERaaa - 1)(o))
4002:                   Next
4003:                   For o As Integer = quantosdciAnosA To A - 1
4004:
4005:                       ordenados(ITERaaa - 1)(o) = "0000000000"
400:                    Next
401:
402:                    'ordenar (é por ordem crescente e o mal é que começa pelos nothing)
                        Array.Sort(ordenados(ITERaaa - 1), 0, A) 'NÃO ADIANTA, ficam só dois no fim na mesma e é o sort que poe mal
403:
                        For z As Integer = quantosdciAnosP To quantosdciAnosA 'diferenca diz quantos (os de menor ranking que) terão ranking = 0
404:                        'os ordenadoss de menos ranking passam a zero
                            'EVITAR QUE STRING NÃO TENHA 6 DIGITOS 
405:                        cruzamentocerto((ordenados(ITERaaa - 1)(z - 1)).ToString.Substring(4, 1)).ranking = "0000000000"
406:

                        Next
407:
                        Array.Reverse(ordenados(ITERaaa - 1), 0, A) 'se A<4 o(s) primeiro(s) fica(m) nulo(s), daí o de 0 a A-1

408:                    For q As Integer = 1 To quantosdciAnosP 'quantos A nos P diz quantos de maior ranking
409:                        'os ordenadoss de maior ranking passam como estão
410:
411:                        'inverter (para ficar por ordem descrescente)
412:                        'EVITAR QUE STRING NÃO TENHA 6 DIGITOS 

414:                        cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp = ITERppp
415:                        cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).RANKING = arraycroces(ITERadorprior).ranking
416:                        iteradorQual1 = ordenados(ITERaaa - 1)(q - 1).ToString.Substring(10, 2) 'cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp & ... 'ERRO 13 = o ordenados não dá número, falta-me o indice
417:                        iteradorQual2 = iteradorQual1
419:                    Next

420:                End If
40:             Else
41:                 'fica como estava
42:             End If
43:         Else
44:             'fica como estava
45:         End If
46:
47:
48:         ITERppp = ITERppp + 1
49:     Loop 'final do while iterPPP

50:     anular(iteradorQual1, iteradorQual2, ITERaaa, arraycroces, cruzamentocerto)
51:     'Else 'se anulado = true
52:     'fica como estava
53:     'End If 'fim da verificação do anulado do iteradorprior

54:     Exit Sub
MOSTRARERRO:
        MsgBox("sub subrotina: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub SUBROTINA2(ByVal ITERppp As Single, ByVal ITERaaa As Single, ByVal ITERadorprior As Single, ByVal iteradorQual1 As SByte, ByVal iteradorQual2 As SByte, ByVal arraycroces As Array, ByVal cruzamentocerto As Array, ByVal quantosdciAnosP As Single, ByVal quantosdciAnosA As Single)
1:      On Error GoTo MOSTRARERRO
        cruzamentocerto(0) = resultado1
        cruzamentocerto(1) = resultado2
        cruzamentocerto(2) = resultado3
        cruzamentocerto(3) = resultado4
2:      Dim ordenados(3) As Object
3:      ordenados(0) = ordenado1
4:      ordenados(1) = ordenado2
5:      ordenados(2) = ordenado3
6:      ordenados(3) = ordenado4


        Dim arraycruza(44) As Object
        arraycruza(11) = cruzamento11
7:      arraycruza(12) = cruzamento12
8:      arraycruza(13) = cruzamento13
9:      arraycruza(14) = cruzamento14
10:     arraycruza(21) = cruzamento21
11:     arraycruza(22) = cruzamento22
12:     arraycruza(23) = cruzamento23
13:     arraycruza(24) = cruzamento24
14:     arraycruza(31) = cruzamento31
15:     arraycruza(32) = cruzamento32
16:     arraycruza(33) = cruzamento33
17:     arraycruza(34) = cruzamento34
18:     arraycruza(41) = cruzamento41
19:     arraycruza(42) = cruzamento42
20:     arraycruza(43) = cruzamento43
21:     arraycruza(44) = cruzamento44


93:     ITERppp = 1
94:     Do While ITERppp <= P
361:
363:        ITERadorprior = Convert.ToSingle((ITERppp) & (ITERaaa))
37:         If arraycroces(ITERadorprior).anulado = False Then
38:             If arraycroces(ITERadorprior).ranking.ToString.Substring(1) >= cruzamentocerto(ITERaaa - 1).ranking.ToString.Substring(1) Then
39:                 If Not (quantosdciAnosP - quantosdciAnosA) < 0 Then
390:                    cruzamentocerto(ITERaaa - 1).qualP = ITERppp
391:                    cruzamentocerto(ITERaaa - 1).ranking = arraycroces(ITERadorprior).ranking
393:                    iteradorQual1 = Convert.ToSByte(cruzamentocerto(ITERaaa - 1).qualP & ITERaaa)
394:                    iteradorQual2 = iteradorQual1
395:                Else         'mais aviados que prescritos 
396:                    'ordeno por ranking
                        ReDim ordenados(ITERaaa - 1)(P - 1)
397:                    For o As Integer = 0 To P - 1 'quantosdciAnosA - 1
                            Dim concata1 = (o + 1) & A
398:                        concatenado = arraycruza(concata1).ranking & concata1
399:                        ordenados(ITERaaa - 1)(o) = concatenado
4001:                       '  MsgBox((ITERaaa - 1) & " " & o & " " & ordenados(ITERaaa - 1)(o))
4002:                   Next
4003:
401:                    'nos dois a seguir era A em vez de P e mudei a 17/12/2013
402:                    Array.Sort(ordenados(ITERaaa - 1), 0, P) 'ordenar (é por ordem crescente)
403:                    Array.Reverse(ordenados(ITERaaa - 1), 0, P) 'fica decrescente

408:                    For q As Integer = 1 To quantosdciAnosP 'quantos A nos P diz quantos de maior ranking
409:                        'os ordenadoss de maior ranking passam como estão
410:                        cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp = ITERppp
411:                        'MsgBox(cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).RANKING)
415:                        cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).RANKING = arraycroces(ITERadorprior).ranking
416:                        iteradorQual1 = ordenados(ITERaaa - 1)(q - 1).ToString.Substring(10, 2) 'cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp & ... 'ERRO 13 = o ordenados não dá número, falta-me o indice
417:                        iteradorQual2 = iteradorQual1
419:                    Next

420:                End If
40:             Else
41:                 'fica como estava
42:             End If
43:         Else
44:             'fica como estava
45:         End If
46:
47:
48:         ITERppp = ITERppp + 1
49:     Loop 'final do while iterPPP

50:     anular(iteradorQual1, iteradorQual2, ITERaaa, arraycroces, cruzamentocerto)


54:     Exit Sub
MOSTRARERRO:
        MsgBox("sub subrotina2: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Sub anular(ByVal iteradorqual1 As SByte, iteradorqual2 As SByte, ByVal iterAAA As Single, ByVal arraycroces As Object, ByVal cruzamentocerto As Object)
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

        iteradorqual1 = iteradorqual1 - 30

50:     Do While iteradorqual1 <= P * 10 + iterAAA
            If iteradorqual1 > 10 Then
51:             arraycroces(iteradorqual1).anulado = True
52:         End If
            iteradorqual1 = iteradorqual1 + 10
53:     Loop
54:
55:     Do While iteradorqual2 <= cruzamentocerto(iterAAA - 1).qualP * 10 + A
            If aceitarduplicados = False Then
56:             arraycroces(iteradorqual2).anulado = True
57:             iteradorqual2 = iteradorqual2 + 1
            End If
58:     Loop
        Exit Sub
MOSTRARERRO:
        MsgBox("sub anular: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub read_only()
        On Error GoTo MOSTRARERRO

        If resultadomostrado = True Then
            For Each t As TextBox In Me.Controls.OfType(Of TextBox)()
                t.ReadOnly = True
                codEC2.ReadOnly = False
                CodeEC1.ReadOnly = False
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



    '    Sub mostrarnocomparador()
    '        On Error GoTo MOSTRARERRO
    '        'entra qualp, indicedoresultado, quais erros
    '        'preenche um com dados do prescrito
    '        'preenche o outro com dados do aviado
    '        'consoante os erros highlita partes de ambos


    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Sub mostrarnocomparador: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub

    Function detectarmaisdedoisaviados() As Boolean
1:      On Error GoTo MOSTRARERRO
2:      Dim aviamento(3) As Object
3:      aviamento(0) = oa1
4:      aviamento(1) = oa2
5:      aviamento(2) = oa3
6:      aviamento(3) = oa4
7:      repeticao = 1
8:      For aviados As Integer = 1 To A - 1
9:          If aviamento(aviados).code = aviamento(aviados - 1).code Then
91:             If (IsNumeric(aviamento(aviados).qty) AndAlso aviamento(aviados).qty <> 1) Or (antesa(aviados) > 0 And antesa(aviados) <> 1) Then
10:                 repeticao = repeticao + 1
                End If
11:             If repeticao = 3 Then
12:                 If detectarmaisdedoisprescritos() Then Return True
13:                 Exit For
14:             End If
15:         End If
16:     Next
17:
18:     Exit Function
MOSTRARERRO:
        MsgBox("Sub detectarmaisdedoisaviados: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Function detectarmaisdedoisprescritos() As Boolean
1:      On Error GoTo MOSTRARERRO
2:      Dim prescricao(3) As Object
3:      prescricao(0) = op1
4:      prescricao(1) = op2
5:      prescricao(2) = op3
6:      prescricao(3) = op4
7:      repeticao = 1
8:      For prescritos As Integer = 1 To P
9:          If prescricao(prescritos).code = prescricao(prescritos - 1).code Then
                If (IsNumeric(prescricao(prescritos).qty) AndAlso prescricao(prescritos).qty <> 1) Or (antesa(prescritos) > 0 And antesa(prescritos) <> 1) Then
10:                 repeticao = repeticao + 1
                End If
11:             If repeticao = 3 Then
12:                 Return True
13:                 Exit For
14:             End If
15:         End If
16:     Next
17:
18:     Exit Function
MOSTRARERRO:
        MsgBox("Sub detectarmaisdeprescritos: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function


    Function PVPmax(actual As Double, menosum As Double, menosdois As Double, menostres As Double, menosquatro As Double, menoscinco As Double) As Double
        On Error GoTo MOSTRARERRO


        Select Case mesactual
            Case Is = 1, 4, 7, 10
                'usam-se 4 (o actual e os 3 anteriores)
                PVPmax = System.Math.Max(actual, (System.Math.Max(menosum, (System.Math.Max(menostres, menosdois)))))
            Case Is = 2, 5, 8, 11
                'usam-se 5 (o actual e os 4 anteriores)
                PVPmax = System.Math.Max(actual, (System.Math.Max(menosum, (System.Math.Max(menostres, System.Math.Max(menosquatro, menosdois))))))
            Case Is = 3, 6, 9, 12
                'usam-se 6 (o actual e os 5 anteriores)
                PVPmax = System.Math.Max(actual, (System.Math.Max(menosum, (System.Math.Max(menostres, System.Math.Max(menosquatro, System.Math.Max(menosdois, menoscinco)))))))
        End Select


        Exit Function
MOSTRARERRO:
        MsgBox("function descobrirPVPmaximo: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        On Error GoTo MOSTRARERRO

        indicar(1)
        somas()
        Exit Sub

MOSTRARERRO: MsgBox("sub ComboBox1_SelectedIndexChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        On Error GoTo MOSTRARERRO

        indicar(2)
        somas()
        Exit Sub

MOSTRARERRO: MsgBox("sub ComboBox2_SelectedIndexChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        On Error GoTo MOSTRARERRO

        indicar(3)
        somas()
        Exit Sub

MOSTRARERRO: MsgBox("sub ComboBox3_SelectedIndexChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub


    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        On Error GoTo MOSTRARERRO

        indicar(4)
        somas()
        Exit Sub

MOSTRARERRO: MsgBox("sub ComboBox4_SelectedIndexChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub




End Class
