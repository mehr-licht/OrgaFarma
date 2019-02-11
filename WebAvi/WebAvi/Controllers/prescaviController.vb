Imports System.Web.Mvc
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
'Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports System.Web.Routing
Imports System.Web.Mvc.Html
Imports System.IO.File
Imports System.IO

Namespace Passdata
    Public Class prescaviController
        Inherits System.Web.Mvc.Controller
        Dim show(3) As embalagem.cruzamento  'assim além do ranking posso ir buscar o qualp e outras coisas

        Dim doaviado(3) As Object
        Dim arrayoa(3) As Object
        Dim arrayop(3) As Object
        Dim arrayarow(3) As Object
        Dim arrayprow(3) As Object
        Dim arrayaviado(3) As Object
        Dim arrayprescrito(3) As Object
        Dim reresult(3) As Object
        Dim resultwebavi1 As New Models.linha
        Dim resultwebavi2 As New Models.linha
        Dim resultwebavi3 As New Models.linha
        Dim resultwebavi4 As New Models.linha
        Dim totalresult As New Models.geral
        Dim arraycombo(3) As Object  'no webavi não se mostra no form logo comenta-se


        Dim arrayvalorcombopvp(3) As Object
        Dim aviam1 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim aviam2 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim aviam3 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim aviam4 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim presc1 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim presc2 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim presc3 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção
        Dim presc4 'não defini estes. serão os parametros que entram. depois está a ser u7tilizada uma propriedade text. atenção

        Dim naoindicar As Boolean


        Dim mesactual As Single = Date.Now.Month
        Dim concatenado As String
        Dim antesp(3) As Double
        Dim depoisp(3) As Double
        Dim antesa(3) As Double
        Dim depoisa(3) As Double
        Dim repeticao As New Integer
        Dim aceitarduplicados As Boolean = False
        Dim verificador As String
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

        Dim filtrolab As Boolean = False

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

        Dim firsttime As Boolean = True
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


        Dim codigoarray As New ArrayList
        Dim codigo41 As New Models.meds
        Dim codigo4 As New Models.meds
        Dim novado1 As Boolean
        Dim novado As Boolean



        Dim p1, p2, p3, p4, a1, a2, a3, a4 As Integer


        Dim descomp1mostrado As Boolean
        Dim descomp2mostrado As Boolean
        Dim descomp3mostrado As Boolean
        Dim descomp4mostrado As Boolean

        Dim Prescrito1 As New Models.meds
        Dim Prescrito2 As New Models.meds
        Dim Prescrito3 As New Models.meds
        Dim Prescrito4 As New Models.meds


        Dim Aviado1 As New Models.meds
        Dim Aviado2 As New Models.meds
        Dim Aviado3 As New Models.meds
        Dim Aviado4 As New Models.meds

        Dim vazio As New Models.meds




        Dim A As Short
        Dim P As Short
        Dim AA As Short
        Dim PP As Short

        Dim procurado1 As Boolean
        Dim procurado2 As Boolean
        Dim procurado3 As Boolean
        Dim procurado4 As Boolean

        'novo webavi
        Dim dadosTA As New bdTableAdapters.dadosTableAdapter
        'Dim infarmedTA As New basededadosDataSetTableAdapters.infarmedTableAdapter

        'novo webavi
        Dim DS As New bd


        'Dim DS As New basededadosDataSet

        Dim row As DataRow




        'novo webavi = faço tudo o que era bas"ededadosDataSet.infarmedRow como bd.dadosRow

        Dim p1row As bd.dadosRow
        Dim p2row As bd.dadosRow
        Dim p3row As bd.dadosRow
        Dim p4row As bd.dadosRow
        Dim a1row As bd.dadosRow
        Dim a2row As bd.dadosRow
        Dim a3row As bd.dadosRow
        Dim a4row As bd.dadosRow



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
        Dim resultadomostrado As Boolean


        <HttpPost>
        Function cnpemfill(pedido) As ActionResult
1:          On Error GoTo MOSTRARERRO

2:              Dim cnpemTA As New bdTableAdapters.dadosTableAdapter
3:              Dim cnpembd As New bd
                If IsNumeric(pedido) Then
4:                  If pedido < 50000000 Then
5:                      pedido = cnpemTA.CnpemByCode(pedido)
6:                  End If
                    If Not pedido = 0 Then
7:                      cnpemTA.FillByCnpem(cnpembd.dados, Convert.ToDecimal(pedido))
8:                      Dim tamanho = cnpembd.Tables(0).Rows.Count
                        If tamanho > 30 Then
                            tamanho = 30
                        End If
9:                      Dim dev((tamanho - 1), 30) As Object

10:                     Dim dd As Integer = 0
11:                     For Each dr As DataRow In cnpembd.dados.Rows
12:                         Dim mm As Integer = 0
13:                         For Each member In dr.ItemArray
14:                             dev(dd, mm) = member 'erro91
15:                             mm = mm + 1
16:                         Next
17:                         dd = dd + 1
18:                     Next
19:
20:                     Dim devdev(tamanho) As Array

21:                     For dv As Integer = 0 To tamanho - 1
22:                         devdev(dv) = {dev(dv, 0), dev(dv, 1), dev(dv, 2), dev(dv, 3), dev(dv, 4), dev(dv, 5), dev(dv, 11)}
25:                     Next

26:                     Return Me.Json(devdev.ToArray, JsonRequestBehavior.AllowGet)
                    Else
                        Return Me.Json("", JsonRequestBehavior.AllowGet)
                    End If
                Else
                    Return Me.Json("", JsonRequestBehavior.AllowGet)
                End If
MOSTRARERRO:
                MsgBox("F16" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
                Return RedirectToAction("Index", "Input")
            Resume Next
        End Function





        <HttpPost>
        Function comparar(presc1 As String, presc2 As String, presc3 As String, presc4 As String, aviad1 As String, aviad2 As String, aviad3 As String, aviad4 As String) As ActionResult
a0:         On Error GoTo MOSTRARERRO
a1:         If presc1 = "" Or presc1 = " " Then
1:              presc1 = "0"
2:          End If
3:          If presc2 = "" Or presc2 = " " Then
4:              presc2 = "0"
5:          End If
6:          If presc3 = "" Or presc3 = " " Then
7:              presc3 = "0"
8:          End If
9:          If presc4 = "" Or presc4 = " " Then
10:             presc4 = "0"
11:         End If
12:         If aviad1 = "" Or aviad1 = " " Then
13:             aviad1 = "0"
14:         End If
15:         If aviad2 = "" Or aviad2 = " " Then
16:             aviad2 = "0"
17:         End If
18:         If aviad3 = "" Or aviad3 = " " Then
19:             aviad3 = "0"
20:         End If
21:         If aviad4 = "" Or aviad4 = " " Then
22:             aviad4 = "0"
23:         End If
24:

25:         dadosTA.FillByCode(DS.dados, Convert.ToDecimal(presc1), Convert.ToDecimal(presc2), Convert.ToDecimal(presc3), Convert.ToDecimal(presc4), Convert.ToDecimal(aviad1), Convert.ToDecimal(aviad2), Convert.ToDecimal(aviad3), Convert.ToDecimal(aviad4))
            '1:          Dim codigos = {presc1, presc2, presc3, presc4, aviad1, aviad2, aviad3, aviad4}
            '2:          For Each codigo In codigos
            '3:              If codigo <> 0 Then
            '4:             'naogravar = DS.dados.FindBycode(codigo).code
            '5:                  Dim teste As bd.dadosRow
            '6:                  teste = DS.dados.FindBycode(codigo)
            '7:                  If IsNothing(teste) Then
            '8:                      gravar(codigo.ToString)
            '9:                  End If

            '10:             End If

            '11:         Next

26:         inicializar(presc1, presc2, presc3, presc4, aviad1, aviad2, aviad3, aviad4)

1901:
1905:       limpo = False
1906:
1907:
            comparado = True
1908:       inicializado = False
1909:
1910:       Dim arraytotal = {totalresult.msg, totalresult.alinea, totalresult.port, totalresult.excep, totalresult.top5, totalresult.cor}
1911:       Dim linha1 = {reresult(0).msg, reresult(0).alinea, reresult(0).port, reresult(0).excep, reresult(0).top5, reresult(0).cor, resultado1.qualP}
32:         Dim linha2 = {reresult(1).msg, reresult(1).alinea, reresult(1).port, reresult(1).excep, reresult(1).top5, reresult(1).cor, resultado2.qualP}
33:         Dim linha3 = {reresult(2).msg, reresult(2).alinea, reresult(2).port, reresult(2).excep, reresult(2).top5, reresult(2).cor, resultado3.qualP}
34:         Dim linha4 = {reresult(3).msg, reresult(3).alinea, reresult(3).port, reresult(3).excep, reresult(3).top5, reresult(3).cor, resultado4.qualP}
35:         Dim mostrarshow1 = {arrayop(0).code, arrayop(0).dci, arrayop(0).nome, arrayop(0).forma, arrayop(0).dose, arrayop(0).qty, arrayop(0).comp, arrayop(0).gh, arrayop(0).pvp, arrayop(0).pr, arrayop(0).gen, arrayop(0).lab, arrayop(0).top5, arrayop(0).pvpmenos1, arrayop(0).pvpmenos2, arrayop(0).CNPEM, arrayop(0).dci_obr}
36:         Dim mostrarshow2 = {arrayop(1).code, arrayop(1).dci, arrayop(1).nome, arrayop(1).forma, arrayop(1).dose, arrayop(1).qty, arrayop(1).comp, arrayop(1).gh, arrayop(1).pvp, arrayop(1).pr, arrayop(1).gen, arrayop(1).lab, arrayop(1).top5, arrayop(1).pvpmenos1, arrayop(1).pvpmenos2, arrayop(1).CNPEM, arrayop(1).dci_obr}
37:         Dim mostrarshow3 = {arrayop(2).code, arrayop(2).dci, arrayop(2).nome, arrayop(2).forma, arrayop(2).dose, arrayop(2).qty, arrayop(2).comp, arrayop(2).gh, arrayop(2).pvp, arrayop(2).pr, arrayop(2).gen, arrayop(2).lab, arrayop(2).top5, arrayop(2).pvpmenos1, arrayop(2).pvpmenos2, arrayop(2).CNPEM, arrayop(2).dci_obr}
38:         Dim mostrarshow4 = {arrayop(3).code, arrayop(3).dci, arrayop(3).nome, arrayop(3).forma, arrayop(3).dose, arrayop(3).qty, arrayop(3).comp, arrayop(3).gh, arrayop(3).pvp, arrayop(3).pr, arrayop(3).gen, arrayop(3).lab, arrayop(3).top5, arrayop(3).pvpmenos1, arrayop(3).pvpmenos2, arrayop(3).CNPEM, arrayop(3).dci_obr}
39:         Dim mostrardoaviado1 = {arrayoa(0).code, arrayoa(0).dci, arrayoa(0).nome, arrayoa(0).forma, arrayoa(0).dose, arrayoa(0).qty, arrayoa(0).comp, arrayoa(0).gh, arrayoa(0).pvp, arrayoa(0).pr, arrayoa(0).gen, arrayoa(0).lab, arrayoa(0).top5, arrayoa(0).pvpmenos1, arrayoa(0).pvpmenos2, arrayoa(0).CNPEM, arrayoa(0).dci_obr, arrayoa(0).pvpmenos3, arrayoa(0).pvpmenos4, arrayoa(0).pvpmenos5, arrayoa(0).d4250, arrayoa(0).d1234, arrayoa(0).d21094, arrayoa(0).d10279, arrayoa(0).d10280, arrayoa(0).d10910, arrayoa(0).d14123, arrayoa(0).lei6}
50:         Dim mostrardoaviado2 = {arrayoa(1).code, arrayoa(1).dci, arrayoa(1).nome, arrayoa(1).forma, arrayoa(1).dose, arrayoa(1).qty, arrayoa(1).comp, arrayoa(1).gh, arrayoa(1).pvp, arrayoa(1).pr, arrayoa(1).gen, arrayoa(1).lab, arrayoa(1).top5, arrayoa(1).pvpmenos1, arrayoa(1).pvpmenos2, arrayoa(1).CNPEM, arrayoa(1).dci_obr, arrayoa(1).pvpmenos3, arrayoa(1).pvpmenos4, arrayoa(1).pvpmenos5, arrayoa(1).d4250, arrayoa(1).d1234, arrayoa(1).d21094, arrayoa(1).d10279, arrayoa(1).d10280, arrayoa(1).d10910, arrayoa(1).d14123, arrayoa(1).lei6}
51:         Dim mostrardoaviado3 = {arrayoa(2).code, arrayoa(2).dci, arrayoa(2).nome, arrayoa(2).forma, arrayoa(2).dose, arrayoa(2).qty, arrayoa(2).comp, arrayoa(2).gh, arrayoa(2).pvp, arrayoa(2).pr, arrayoa(2).gen, arrayoa(2).lab, arrayoa(2).top5, arrayoa(2).pvpmenos1, arrayoa(2).pvpmenos2, arrayoa(2).CNPEM, arrayoa(2).dci_obr, arrayoa(2).pvpmenos3, arrayoa(2).pvpmenos4, arrayoa(2).pvpmenos5, arrayoa(2).d4250, arrayoa(2).d1234, arrayoa(2).d21094, arrayoa(2).d10279, arrayoa(2).d10280, arrayoa(2).d10910, arrayoa(2).d14123, arrayoa(2).lei6}
52:         Dim mostrardoaviado4 = {arrayoa(3).code, arrayoa(3).dci, arrayoa(3).nome, arrayoa(3).forma, arrayoa(3).dose, arrayoa(3).qty, arrayoa(3).comp, arrayoa(3).gh, arrayoa(3).pvp, arrayoa(3).pr, arrayoa(3).gen, arrayoa(3).lab, arrayoa(3).top5, arrayoa(3).pvpmenos1, arrayoa(3).pvpmenos2, arrayoa(3).CNPEM, arrayoa(3).dci_obr, arrayoa(3).pvpmenos3, arrayoa(3).pvpmenos4, arrayoa(3).pvpmenos5, arrayoa(3).d4250, arrayoa(3).d1234, arrayoa(3).d21094, arrayoa(3).d10279, arrayoa(3).d10280, arrayoa(3).d10910, arrayoa(3).d14123, arrayoa(3).lei6}

            Return mostrar(arraytotal, linha1, linha2, linha3, linha4, mostrarshow1, mostrarshow2, mostrarshow3, mostrarshow4, mostrardoaviado1, mostrardoaviado2, mostrardoaviado3, mostrardoaviado4)
            Exit Function
MOSTRARERRO:
            MsgBox("F01" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function


        '        Function gravar(texto)
        '            On Error GoTo MOSTRARERRO

        '1:          Dim path = ("../log.txt")
        '            If System.IO.File.Exists(path) Then
        '                'do nothing
        '                MsgBox("existe")
        '            Else
        '                System.IO.File.Create(path)
        '                MsgBox("criou")
        '            End If

        '            Dim tw As New StreamWriter(path)
        '            tw.Write(texto)
        '9:          tw.Close()
        '            MsgBox("gravou")
        '            '12:         Using sw As StreamWriter = System.IO.File.AppendText(path)
        '            '                MsgBox("vaigravar")
        '            '13:             sw.WriteLine(texto)
        '            '14:             sw.Close()
        '            '                MsgBox("gravou")
        '            '15:         End Using




        '            Exit Function
        'MOSTRARERRO:
        '            MsgBox("F18" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
        '            Resume Next
        '        End Function

        Function inicializar(ByVal inicP1 As String, ByVal inicP2 As String, ByVal inicP3 As String, ByVal inicP4 As String, ByVal inicA1 As String, ByVal inicA2 As String, ByVal inicA3 As String, ByVal inicA4 As String)

            On Error GoTo MOSTRARERRO

            naoindicar = True

11127:      concatenado = "00"
11128:      repeticao = 0
11129:      Array.Clear(antesp, 0, 4)
11130:      Array.Clear(antesa, 0, 4)
11131:      Array.Clear(depoisp, 0, 4)
11132:      Array.Clear(depoisa, 0, 4)

            'webavi até aqui ok

            'webavi
            '11133:      C_PVP_1.Text = ""
            '11134:      C_PVP_2.Text = ""
            '11135:      C_PVP_3.Text = ""
            '11136:      C_PVP_4.Text = ""
            '11137:      C_PVP_1.BackColor = SystemColors.GradientInactiveCaption
            '11138:      C_PVP_2.BackColor = SystemColors.GradientInactiveCaption
            '11139:      C_PVP_3.BackColor = SystemColors.GradientInactiveCaption
            '11140:      C_PVP_4.BackColor = SystemColors.GradientInactiveCaption
            '11141:      top5Box.Text = ""
            '11142:      excepBox.Text = ""
11143:      resultadomostrado = False
11144:
            'webavi
            'read_only()
11145:      verificador = "green"
11146:      erro = False
11147:      comparado = False





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
            'webavi até aqui ok


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
            'sehaport.Text = ""
            limpo = True
            agrupado = False



a:          varlabelcruz1 = ""
b:         ' labelcruz1.Text = ""
c:          varlabelcruz2 = ""
d:         ' labelcruz2.Text = ""
e:          varlabelcruz3 = ""
f:         ' labelcruz3.Text = ""
g:          varlabelcruz4 = ""
h:
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


72:         Prescrito1.codigo = Nothing
73:         Prescrito2.codigo = Nothing
74:         Prescrito3.codigo = Nothing
75:         Prescrito4.codigo = Nothing
76:         Aviado1.codigo = Nothing
77:         Aviado2.codigo = Nothing
78:         Aviado3.codigo = Nothing
79:         Aviado4.codigo = Nothing
80:         a1p1.nivel = 100
81:         a2p1.nivel = 100
82:         a3p1.nivel = 100
83:         a4p1.nivel = 100
84:         a1p2.nivel = 100
85:         a2p2.nivel = 100
86:         a3p2.nivel = 100
87:         a4p2.nivel = 100
88:         a1p3.nivel = 100
89:         a2p3.nivel = 100
90:         a3p3.nivel = 100
91:         a4p3.nivel = 100
92:         a1p4.nivel = 100
93:         a2p4.nivel = 100
94:         a3p4.nivel = 100
95:         a4p4.nivel = 100
96:         a1p1.resultado = 100
97:         a2p1.resultado = 100
98:         a3p1.resultado = 100
99:         a4p1.resultado = 100
100:        a1p2.resultado = 100
101:        a2p2.resultado = 100
102:        a3p2.resultado = 100
103:        a4p2.resultado = 100
104:        a1p3.resultado = 100
105:        a2p3.resultado = 100
106:        a3p3.resultado = 100
107:        a4p3.resultado = 100
108:        a1p4.resultado = 100
109:        a2p4.resultado = 100
110:        a3p4.resultado = 100
111:        a4p4.resultado = 100
112:        a1p1.mostrado = False
113:        a2p1.mostrado = False
114:        a3p1.mostrado = False
115:        a4p1.mostrado = False
116:        a1p2.mostrado = False
117:        a2p2.mostrado = False
118:        a3p2.mostrado = False
119:        a4p2.mostrado = False
120:        a1p3.mostrado = False
121:        a2p3.mostrado = False
122:        a3p3.mostrado = False
123:        a4p3.mostrado = False
124:        a1p4.mostrado = False
125:        a2p4.mostrado = False
126:        a3p4.mostrado = False
127:        a4p4.mostrado = False
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
156:
158:        varcruzp1 = 0
159:        varcruzp2 = 0
160:        varcruzp3 = 0
161:        varcruzp4 = 0
162:

195:        verde = True
196:        amarelo = False
197:        vermelho = False
            ''198:        verifgenlabel.Text = ""
            ''199:        verifgenlabel.BackColor = Color.Transparent
19901:      conjunto = 0
19902:
19903:      Select Case mesactual
                Case Is = 1, 4, 7, 10
19905:          'usam-se 4 (o actual e os 3 anteriores)
19906:              oa1.arraypvp = {"", "", "", ""}
19907:              oa2.arraypvp = {"", "", "", ""}
19908:              oa3.arraypvp = {"", "", "", ""}
19909:              oa4.arraypvp = {"", "", "", ""}
19910:          Case Is = 2, 5, 8, 11
19911:          'usam-se 5 (o actual e os 4 anteriores)
19912:              oa1.arraypvp = {"", "", "", "", ""}
19913:              oa2.arraypvp = {"", "", "", "", ""}
19914:              oa3.arraypvp = {"", "", "", "", ""}
19915:              oa4.arraypvp = {"", "", "", "", ""}
19916:          Case Is = 3, 6, 9, 12
19917:          'usam-se 6 (o actual e os 5 anteriores)
19918:              oa1.arraypvp = {"", "", "", "", "", ""}
19919:              oa2.arraypvp = {"", "", "", "", "", ""}
19920:              oa3.arraypvp = {"", "", "", "", "", ""}
19921:              oa4.arraypvp = {"", "", "", "", "", ""}
19922:      End Select
19923:      Select Case mesactual
                Case Is = 1, 4, 7, 10
19925:          'usam-se 4 (o actual e os 3 anteriores)
19926:              Array.Clear(oa1.arraypvp, 0, 4)
19927:              Array.Clear(oa2.arraypvp, 0, 4)
19928:              Array.Clear(oa3.arraypvp, 0, 4)
19929:              Array.Clear(oa4.arraypvp, 0, 4)
19930:          Case Is = 2, 5, 8, 11
19931:          'usam-se 5 (o actual e os 4 anteriores)
19932:              Array.Clear(oa1.arraypvp, 0, 5)
19933:              Array.Clear(oa2.arraypvp, 0, 5)
19934:              Array.Clear(oa3.arraypvp, 0, 5)
19935:              Array.Clear(oa4.arraypvp, 0, 5)
19936:          Case Is = 3, 6, 9, 12
19937:          'usam-se 6 (o actual e os 5 anteriores)
19938:              Array.Clear(oa1.arraypvp, 0, 6)
19939:              Array.Clear(oa2.arraypvp, 0, 6)
19940:              Array.Clear(oa3.arraypvp, 0, 6)
19941:              Array.Clear(oa4.arraypvp, 0, 6)
19942:      End Select
19943:
19944:
19945:
19952:
19953:      firsttime = False
19954:      naoindicar = False


            'webavi
            irbuscar2013(inicP1, inicP2, inicP3, inicP4, inicA1, inicA2, inicA3, inicA4)


            Exit Function
MOSTRARERRO:
            MsgBox("F02" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function





        Function irbuscar2013(ByVal irbuscP1 As String, ByVal irbuscP2 As String, ByVal irbuscP3 As String, ByVal irbuscP4 As String, ByVal irbuscA1 As String, ByVal irbuscA2 As String, ByVal irbuscA3 As String, ByVal irbuscA4 As String)
1:          On Error GoTo MOSTRARERRO
2:

325:        Select Case Len(irbuscP1)
                Case Is = 7
327:                op1.porCNPEM = False
328:            Case Is = 8
329:                op1.porCNPEM = True
330:            Case Else
331:                irbuscP1 = "0"
332:        End Select
333:        Select Case Len(irbuscP2)
                Case Is = 7
335:                op2.porCNPEM = False
336:            Case Is = 8
337:                op2.porCNPEM = True
338:            Case Else
339:                irbuscP2 = "0"
340:        End Select
341:        Select Case Len(irbuscP3)
                Case Is = 7
343:                op3.porCNPEM = False
344:            Case Is = 8
345:                op3.porCNPEM = True
346:            Case Else
347:                irbuscP3 = "0"
348:        End Select
349:        Select Case Len(irbuscP4)
                Case Is = 7
351:                op4.porCNPEM = False
352:            Case Is = 8
353:                op4.porCNPEM = True
354:            Case Else
355:                irbuscP4 = "0"
356:        End Select
357:        Select Case Len(irbuscA1)
                Case Is = 7
359:                oa1.porCNPEM = False
360:            Case Else
361:                irbuscA1 = "0"
362:        End Select
363:        Select Case Len(irbuscA2)
                Case Is = 7
365:                oa2.porCNPEM = False
366:
367:            Case Else
368:                irbuscA2 = "0"
369:        End Select
370:        Select Case Len(irbuscA3)
                Case Is = 7
372:                oa3.porCNPEM = False
373:
374:            Case Else
375:                irbuscA3 = "0"
376:        End Select
377:        Select Case Len(irbuscA4)
                Case Is = 7
379:                oa4.porCNPEM = False
380:
381:            Case Else
382:                irbuscA4 = "0"
383:        End Select




303:        Aviado1.codigo = irbuscA1
304:        Aviado2.codigo = irbuscA2
305:        Aviado3.codigo = irbuscA3
306:        Aviado4.codigo = irbuscA4
307:        Prescrito1.codigo = irbuscP1
308:        Prescrito2.codigo = irbuscP2
309:        Prescrito3.codigo = irbuscP3
310:        Prescrito4.codigo = irbuscP4
311:        oa1.code = irbuscA1
312:        oa2.code = irbuscA2
313:        oa3.code = irbuscA3
314:        oa4.code = irbuscA4
315:        op1.code = irbuscP1
316:        op2.code = irbuscP2
317:        op3.code = irbuscP3
318:        op4.code = irbuscP4

            a1row = DS.dados.FindBycode(Aviado1.codigo)

320:        a2row = DS.dados.FindBycode(Aviado2.codigo)
321:        a3row = DS.dados.FindBycode(Aviado3.codigo)
322:        a4row = DS.dados.FindBycode(Aviado4.codigo)
323:
324:

384:
385:
386:
3:          If op4.code > 0 Then
4:              PP = 4
5:          ElseIf op3.code > 0 Then
6:              PP = 3
7:          ElseIf op2.code > 0 Then
8:              PP = 2
9:          Else : PP = 1
10:         End If
11:
12:         If oa4.code > 0 Then
13:             AA = 4
14:         ElseIf oa3.code > 0 Then
15:             AA = 3
16:         ElseIf oa2.code > 0 Then
17:             AA = 2
18:         Else : AA = 1
19:         End If
20:
21:
22:
23:         If irbuscA4 <> "0" And irbuscA4 <> "" Then
24:             A = 4
25:             If irbuscA4 = " " Then
26:                 A = 3
27:             End If
28:         ElseIf irbuscA3 <> "0" And irbuscA3 <> "" Then
29:             A = 3
30:             If irbuscA3 = " " Then
31:                 A = 2
32:             End If
33:         ElseIf irbuscA2 <> "0" And irbuscA2 <> "" Then
34:             A = 2
35:             If irbuscA2 = " " Then
36:                 A = 1
37:             End If
38:
39:         Else : A = 1
40:         End If
41:
42:         If irbuscP4 <> "0" And irbuscP4 <> "" Then
43:             P = 4
44:             If irbuscP4 = " " Then
45:                 P = 3
46:             End If
47:         ElseIf irbuscP3 <> "0" And irbuscP3 <> "" Then
48:             P = 3
49:             If irbuscP3 = " " Then
50:                 P = 2
51:             End If
52:         ElseIf irbuscP2 <> "0" And irbuscP2 <> "" Then
53:             P = 2
54:             If irbuscP2 = " " Then
55:                 P = 1
56:             End If
57:         Else : P = 1
58:         End If
59:         'webavi   dá 0  e nada    MsgBox("forma: " & oa1.forma & vbCr & "principio: " & Aviado1.principio)

592:        arrayoa(0) = oa1
593:        arrayoa(1) = oa2
594:        arrayoa(2) = oa3
595:        arrayoa(3) = oa4

597:        arrayop(0) = op1
598:        arrayop(1) = op2
599:        arrayop(2) = op3
600:        arrayop(3) = op4

602:        arrayarow(0) = a1row
603:        arrayarow(1) = a2row
604:        arrayarow(2) = a3row
605:        arrayarow(3) = a4row

607:        arrayprow(0) = p1row
608:        arrayprow(1) = p2row
609:        arrayprow(2) = p3row
610:        arrayprow(3) = p4row



622:        arrayaviado(0) = Aviado1
623:        arrayaviado(1) = Aviado2
624:        arrayaviado(2) = Aviado3
625:        arrayaviado(3) = Aviado4

627:        arrayprescrito(0) = Prescrito1
628:        arrayprescrito(1) = Prescrito2
629:        arrayprescrito(2) = Prescrito3
630:        arrayprescrito(3) = Prescrito4
631:



632:        For letraA As Integer = 0 To A - 1
                'webavi - substituí
6321:           arrayarow(letraA) = DS.dados.FindBycode(arrayaviado(letraA).codigo)
633:            'arrayarow(letraA) = DS.infarmed.FindBycode(arrayaviado(letraA).codigo)
634:
                'webavi
                'arrayaarray(letraA).Add(arrayarow(letraA))
635:            If Not IsNothing(arrayarow(letraA)) Then
636:                arrayoa(letraA).code = arrayarow(letraA)(0)
637:                arrayoa(letraA).nome = arrayarow(letraA)(1) 'webavi era 2
638:                arrayoa(letraA).d10910 = arrayarow(letraA)(26) 'webavi era 15
639:                arrayoa(letraA).pvpmenos2 = arrayarow(letraA)(15) 'webavi era 21
640:                arrayoa(letraA).cnpem = arrayarow(letraA)(16) 'webavi era 22
641:                arrayoa(letraA).dci = LCase(arrayarow(letraA)(2)) 'webavi era 1
70:                 arrayoa(letraA).forma = arrayarow(letraA)(3)
71:                 arrayoa(letraA).dose = arrayarow(letraA)(4)
72:                 arrayoa(letraA).qty = arrayarow(letraA)(5).ToString.Replace(".", ",")
73:                 arrayoa(letraA).comp = arrayarow(letraA)(6)
74:                 arrayoa(letraA).gh = arrayarow(letraA)(7)
75:                 arrayoa(letraA).gen = arrayarow(letraA)(10) 'webavi era 8
76:                 arrayoa(letraA).lab = arrayarow(letraA)(11) 'webavi era 9
77:                 arrayoa(letraA).pvp = arrayarow(letraA)(8) 'webavi era 23
78:                 arrayoa(letraA).pr = arrayarow(letraA)(9) 'webavi era 24
79:                 arrayoa(letraA).d4250 = arrayarow(letraA)(21) 'webavi era 10
80:                 arrayoa(letraA).d1234 = arrayarow(letraA)(22) 'webavi era 11
81:                 arrayoa(letraA).d10279 = arrayarow(letraA)(24) 'webavi era 13
82:                 arrayoa(letraA).d10280 = arrayarow(letraA)(25) 'webavi era 14
83:                 arrayoa(letraA).d21094 = arrayarow(letraA)(23) 'webavi era 12
84:                 arrayoa(letraA).lei6 = arrayarow(letraA)(28) 'webavi era 19

85:                 arrayoa(letraA).d14123 = arrayarow(letraA)(27) 'webavi era 16
86:                 arrayoa(letraA).dci_obr = arrayarow(letraA)(13) 'webavi era 18
87:                 'arrayoa(letraA).trocamarca = arrayarow(letraA)(20) 'webavi era 20
88:                 arrayoa(letraA).top5 = arrayarow(letraA)(12) 'webavi era 17
89:                 arrayoa(letraA).pvpmenos1 = arrayarow(letraA)(14) 'webavi era 25
8911:               arrayoa(letraA).pvpmenos3 = arrayarow(letraA)(17) 'webavi era 26
8912:               arrayoa(letraA).pvpmenos4 = arrayarow(letraA)(19) 'webavi era 27
8913:               arrayoa(letraA).pvpmenos5 = arrayarow(letraA)(20) 'webavi era 28
8914:
8915:               Select Case mesactual
                        Case Is = 1, 4, 7, 10
8917:                  'usam-se 4 (o actual e os 3 anteriores)
8918:                       ReDim arrayoa(letraA).arraypvp(3)
8919:                       arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3}
8920:                   Case Is = 2, 5, 8, 11
8921:                  'usam-se 5 (o actual e os 4 anteriores)
8922:                       ReDim arrayoa(letraA).arraypvp(4)
8923:                       arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3, arrayoa(letraA).pvpmenos4}
8924:
8925:                   Case Is = 3, 6, 9, 12
8926:                  'usam-se 6 (o actual e os 5 anteriores)
8927:                       ReDim arrayoa(letraA).arraypvp(5)
8928:                       arrayoa(letraA).arraypvp = {arrayoa(letraA).pvp, arrayoa(letraA).pvpmenos1, arrayoa(letraA).pvpmenos2, arrayoa(letraA).pvpmenos3, arrayoa(letraA).pvpmenos4, arrayoa(letraA).pvpmenos5}
8929:
8930:               End Select

                    'webavi
                    '8931:               arraycombo(letraA).DataSource = arrayoa(letraA).arraypvp
                    '8932:               arrayvalorcombopvp(letraA) = True
                    '8933:               valorcombopvp1 = arrayvalorcombopvp(0) 'senão não  funciona não sei porquê
                    '8934:               valorcombopvp2 = arrayvalorcombopvp(1) 'senão não  funciona não sei porquê
                    '8935:               valorcombopvp3 = arrayvalorcombopvp(2) 'senão não  funciona não sei porquê
                    '8936:               valorcombopvp4 = arrayvalorcombopvp(3) 'senão não  funciona não sei porquê
8937:
8938:       'no webavi não se mostra no form logo comenta-se
8939:                   'ComboBox1.DataSource = oa1.arraypvp
8940:                'ComboBox2.DataSource = oa2.arraypvp
8941:                   'ComboBox3.DataSource = oa3.arraypvp
8942:                   'ComboBox4.DataSource = oa4.arraypvp
8943:
8944:               If IsNumeric(arrayoa(letraA).qty) Then
892:                    arrayoa(letraA).pvpun = arrayoa(letraA).pvp / arrayoa(letraA).qty
893:                    arrayoa(letraA).pvpmenos1un = arrayoa(letraA).pvpmenos1 / arrayoa(letraA).qty
894:                    arrayoa(letraA).pvpmenos2un = arrayoa(letraA).pvpmenos2 / arrayoa(letraA).qty
895:                Else
896:                    arrayoa(letraA).pvpun = arrayoa(letraA).pvp
897:                    arrayoa(letraA).pvpmenos1un = arrayoa(letraA).pvpmenos1
898:                    arrayoa(letraA).pvpmenos2un = arrayoa(letraA).pvpmenos2
899:                End If
900:            End If
901:        Next
902:
903:
904:        For letraP As Integer = 0 To P - 1
                'webavi - substituí
                arrayprow(letraP) = DS.dados.FindBycode(arrayprescrito(letraP).codigo)
                '    arrayprow(letraP) = DS.infarmed.FindBycode(arrayprescrito(letraP).code)
906:           ' MsgBox(arrayprescrito(letraP).codigo & vbCr & arrayprow(letraP)(0)) 'dáerro no webavi
                'webavi
                'arrayparray(letraP).Add(arrayprow(letraP))
907:            If Not IsNothing(arrayprow(letraP)) Then
908:                arrayop(letraP).code = arrayprow(letraP)(0)
909:                arrayop(letraP).nome = arrayprow(letraP)(1) 'webavi era 2
910:                arrayop(letraP).d10910 = arrayprow(letraP)(26) 'webavi era 15
911:                arrayop(letraP).pvpmenos2 = arrayprow(letraP)(15) 'webavi era 21
912:                arrayop(letraP).cnpem = arrayprow(letraP)(16) 'webavi era 22
913:                arrayop(letraP).dci = LCase(arrayprow(letraP)(2)) 'webavi era 1
914:                arrayop(letraP).forma = arrayprow(letraP)(3)
915:                arrayop(letraP).dose = arrayprow(letraP)(4)
916:                arrayop(letraP).qty = arrayprow(letraP)(5).ToString.Replace(".", ",")
917:                arrayop(letraP).comp = arrayprow(letraP)(6)
918:                arrayop(letraP).gh = arrayprow(letraP)(7)
919:                arrayop(letraP).gen = arrayprow(letraP)(10) 'webavi era 8
920:                arrayop(letraP).lab = arrayprow(letraP)(11) 'webavi era 9
921:                arrayop(letraP).pvp = arrayprow(letraP)(8) 'webavi era 23
922:                arrayop(letraP).pr = arrayprow(letraP)(9) 'webavi era 24
923:                arrayop(letraP).d4250 = arrayprow(letraP)(21) 'webavi era 10
924:                arrayop(letraP).d1234 = arrayprow(letraP)(22) 'webavi era 11
925:                arrayop(letraP).d10279 = arrayprow(letraP)(24) 'webavi era 13
926:                arrayop(letraP).d10280 = arrayprow(letraP)(25) 'webavi era 14
927:                arrayop(letraP).d21094 = arrayprow(letraP)(23) 'webavi era 12
928:                arrayop(letraP).lei6 = arrayprow(letraP)(28) 'webavi era 19
929:                arrayop(letraP).d14123 = arrayprow(letraP)(27) 'webavi era 16
930:                arrayop(letraP).dci_obr = arrayprow(letraP)(13) 'webavi era 18
931:                 'arrayop(letraP).trocamarca = arrayprow(letraP)(20) 'webavi era 20
932:                arrayop(letraP).top5 = arrayprow(letraP)(12) 'webavi era 17
933:                arrayop(letraP).pvpmenos1 = arrayprow(letraP)(14) 'webavi era 25
934:                arrayop(letraP).pvpmenos3 = arrayprow(letraP)(17) 'webavi era 26
935:                arrayop(letraP).pvpmenos4 = arrayprow(letraP)(19) 'webavi era 27
936:                arrayop(letraP).pvpmenos5 = arrayprow(letraP)(20) 'webavi era 28
937:                arrayop(letraP).pvpexcepC = PVPmax(arrayop(letraP).pvp, arrayop(letraP).pvpmenos1, arrayop(letraP).pvpmenos2, arrayop(letraP).pvpmenos3, arrayop(letraP).pvpmenos4, arrayop(letraP).pvpmenos5)
190:                If IsNumeric(arrayop(letraP).qty) Then
191:                    arrayop(letraP).pvpun = arrayop(letraP).pvp / arrayop(letraP).qty
192:                    arrayop(letraP).pvpmenos1un = arrayop(letraP).pvpmenos1 / arrayop(letraP).qty
193:                    arrayop(letraP).pvpmenos2un = arrayop(letraP).pvpmenos2 / arrayop(letraP).qty
194:                Else
195:                    arrayop(letraP).pvpun = arrayop(letraP).pvp
196:                    arrayop(letraP).pvpmenos1un = arrayop(letraP).pvpmenos1
197:                    arrayop(letraP).pvpmenos2un = arrayop(letraP).pvpmenos2
198:                End If
199:            End If
200:        Next
201:
233:        comparador2013(arrayop, arrayoa)

234:        Exit Function
MOSTRARERRO:
            MsgBox("F03" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function



        Function detectarmaisdedoisaviados(ByVal oa1 As Object, ByVal oa2 As Object, ByVal oa3 As Object, ByVal oa4 As Object) As Boolean
1:          On Error GoTo MOSTRARERRO
2:          Dim aviamento(3) As Object
3:          aviamento(0) = oa1
4:          aviamento(1) = oa2
5:          aviamento(2) = oa3
6:          aviamento(3) = oa4
7:          repeticao = 1
8:          For aviados As Integer = 1 To A - 1
9:              If aviamento(aviados).code = aviamento(aviados - 1).code Then
91:                 If (IsNumeric(aviamento(aviados).qty) AndAlso aviamento(aviados).qty <> 1) Or (antesa(aviados) > 0 And antesa(aviados) <> 1) Then
10:                     repeticao = repeticao + 1
                    End If
11:                 If repeticao = 3 Then
12:                     If detectarmaisdedoisprescritos(op1, op2, op3, op4) Then Return True
13:                     Exit For
14:                 End If
15:             End If
16:         Next
17:
18:         Exit Function
MOSTRARERRO:
            MsgBox("F04" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

            Resume Next
        End Function

        Function detectarmaisdedoisprescritos(ByVal op1 As Object, ByVal op2 As Object, ByVal op3 As Object, ByVal op4 As Object) As Boolean
1:          On Error GoTo MOSTRARERRO
2:          Dim prescricao(3) As Object
3:          prescricao(0) = op1
4:          prescricao(1) = op2
5:          prescricao(2) = op3
6:          prescricao(3) = op4
7:          repeticao = 1
8:          For prescritos As Integer = 1 To P
9:              If prescricao(prescritos).code = prescricao(prescritos - 1).code Then
                    If (IsNumeric(prescricao(prescritos).qty) AndAlso prescricao(prescritos).qty <> 1) Or (antesa(prescritos) > 0 And antesa(prescritos) <> 1) Then
10:                     repeticao = repeticao + 1
                    End If
11:                 If repeticao = 3 Then
12:                     Return True
13:                     Exit For
14:                 End If
15:             End If
16:         Next
17:
18:         Exit Function
MOSTRARERRO:
            MsgBox("F05" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

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
            MsgBox("F06" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

            Resume Next
        End Function


        Function novoprioridade2013()
1:          On Error GoTo MOSTRARERRO
2:
3:          Dim iterPPP As Single = 1
4:          Dim iterAAA As Single = 1
5:          Dim arraycroces(45) As Object
6:          Dim concatP1 As String
7:          Dim concatP2 As String
8:          Dim concatP3 As String
9:          Dim concatP4 As String
10:         Dim concatA1 As String
11:         Dim concatA2 As String
12:         Dim concatA3 As String
13:         Dim concatA4 As String
14:         concatP1 = 0
15:         concatP2 = 0
16:         concatP3 = 0
17:         concatP4 = 0
18:         concatA1 = 0
19:         concatA2 = 0
20:         concatA3 = 0
21:         concatA4 = 0
26:         arraycroces(11) = cruzamento11
27:         arraycroces(12) = cruzamento12
28:         arraycroces(13) = cruzamento13
29:         arraycroces(14) = cruzamento14
30:         arraycroces(21) = cruzamento21
31:         arraycroces(22) = cruzamento22
32:         arraycroces(23) = cruzamento23
33:         arraycroces(24) = cruzamento24
34:         arraycroces(31) = cruzamento31
35:         arraycroces(32) = cruzamento32
36:         arraycroces(33) = cruzamento33
37:         arraycroces(34) = cruzamento34
38:         arraycroces(41) = cruzamento41
39:         arraycroces(42) = cruzamento42
40:         arraycroces(43) = cruzamento43
41:         arraycroces(44) = cruzamento44
42:
46:         Dim cruzamentocerto(3) As Object
47:         cruzamentocerto(0) = resultado1
48:         cruzamentocerto(1) = resultado2
49:         cruzamentocerto(2) = resultado3
50:         cruzamentocerto(3) = resultado4

51:         Dim iteradorPrior As Single = 11
52:         Dim iteradorQual1 As SByte = 11
53:         Dim iteradorQual2 As SByte = 11
54:
55:         Dim arraydciP(3) As String
56:         Dim arraydciA(3) As String
57:
58:
70:
72:         Dim arrayoa(3) As Object
73:         arrayoa(0) = oa1
74:         arrayoa(1) = oa2
75:         arrayoa(2) = oa3
76:         arrayoa(3) = oa4
77:         Dim arrayop(3) As Object
78:         arrayop(0) = op1
79:         arrayop(1) = op2
80:         arrayop(2) = op3
81:         arrayop(3) = op4
82:         Dim quantosdciAnosP(3) As Single
83:         quantosdciAnosP(0) = quantosDCIa1nosP
84:         quantosdciAnosP(1) = quantosDCIa2nosP
85:         quantosdciAnosP(2) = quantosDCIa3nosP
86:         quantosdciAnosP(3) = quantosDCIa4nosP
87:         Dim quantosdciAnosA(3) As Single
88:         quantosdciAnosA(0) = quantosDCIa1
89:         quantosdciAnosA(1) = quantosDCIa2
90:         quantosdciAnosA(2) = quantosDCIa3
91:         quantosdciAnosA(3) = quantosDCIa4
92:
93:         Dim concatP(3) As String
94:         concatP(0) = concatP1
95:         concatP(1) = concatP2
96:         concatP(2) = concatP3
97:         concatP(3) = concatP4
98:         Dim concatA(3) As String
99:         concatA(0) = concatA1
100:        concatA(1) = concatA2
101:        concatA(2) = concatA3
102:        concatA(3) = concatA4
103:
104:        For n As Integer = 1 To P
105:            arraydciP(n - 1) = arrayop(n - 1).dci
106:        Next
107:        For nn As Integer = 1 To A
108:            arraydciA(nn - 1) = arrayoa(nn - 1).dci
109:        Next

            For nnn As Integer = 1 To A
                For nnnn As Integer = 1 To P
110:                If arraydciP(nnnn - 1) = arrayoa(nnn - 1).dci Then
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




124:        Do While iterAAA <= A
125:
126:            Select Case quantosdciAnosP(iterAAA - 1)
                    Case Is = 0
128:                    cruzamentocerto(iterAAA - 1).ranking = "0000000000" 'AVIADO NÃO PRESCRiTO
129:                'Case Is = 1 'só um prescrito, ligação directa 'pus plickas para o caso de haver mais aviados e assim vai à regra geral
130:                '   cruzamentocerto(iterAAA - 1).qualP = iterPPP
131:                '  cruzamentocerto(iterAAA - 1).ranking = arraycroces(iteradorPrior).ranking
132:                ' iteradorQual1 = Convert.ToSByte(cruzamentocerto(iterAAA - 1).qualP & iterAAA)
133:                'anular(iteradorQual1, iteradorQual2, iterAAA, arraycroces, cruzamentocerto)

134:                Case Else
135:                    SUBROTINA2(iterPPP, iterAAA, iteradorPrior, iteradorQual1, iteradorQual2, arraycroces, cruzamentocerto, quantosdciAnosP(iterAAA - 1), quantosdciAnosA(iterAAA - 1))
136:            End Select
137:

146:
                If aceitarduplicados = True Then
150:            'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
165:                For AAA As Integer = 1 To A
169:                    If AAA <> iterAAA Then
                            If (arrayoa(AAA - 1).cnpem > 0 And (arrayoa(AAA - 1).cnpem = arrayoa(iterAAA - 1).cnpem)) Or (arrayoa(AAA - 1).code = arrayoa(iterAAA - 1).code) Then
170:                            If iterAAA > AAA Then 'para não fazer duas vezes e desfazer o que estava feito
                                    cruzamentocerto(iterAAA - 1).ranking = cruzamentocerto(AAA - 1).ranking
                                End If
172:                        End If
173:                    End If
174:                Next
                End If

151:
152:        'MsgBox(iterAAA - 1 & " " & cruzamentocerto(iterAAA - 1).ranking)
159:            iterAAA = iterAAA + 1
160:        Loop 'final do while iterAAA

1601:       showresults()
161:    'para a CC podem ser aviados mais que os prescritos
162:    'COLOCAR a verificação de mesmo dci/forma/dose entre vários aviados (chamada da sub)
            'ou aqui ou dentro do loop ou no fim da subrotina (<= deve ter melhor acesso a todas as variaveis)
            'se mesmo dci/forma/dose e há menos prescritos nessas condições e desnivel de 1,5x entre total qty prescrita e cada um desses aviados
            'então soma quantidade aviada desses prescritos e compara com a quantidade total prescrita.
            'se <= 1,5 => OK para 3808 e necessita justificação para mim (e)
            'else erro (h)
163:        Exit Function
MOSTRARERRO:
            MsgBox("F07" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function


        Function showresults()

1:
            'webavi
            ' On Error GoTo MOSTRARERRO
2:      '    Dim show(3) As embalagem.cruzamento  'assim além do ranking posso ir buscar o qualp e outras coisas
3:          show(0) = resultado1
4:          show(1) = resultado2
5:          show(2) = resultado3
6:          show(3) = resultado4

7:            'webavi  
            'Dim reresult(3) As Object 'era a label que mostrava os resultados passou a ser array para ir buscar ao mostrar resultado no webavi e inicializada no inicio
8:          reresult(0) = resultwebavi1
9:          reresult(1) = resultwebavi2
10:         reresult(2) = resultwebavi3
11:         reresult(3) = resultwebavi4

            '9:          reresult(1) = result2
            '10:         reresult(2) = result3
            '11:         reresult(3) = result4
            ' Dim doaviado(3) As Object
12:         doaviado(0) = oa1
13:         doaviado(1) = oa2
14:         doaviado(2) = oa3
15:         doaviado(3) = oa4

16:         Dim doprescrito(3) As Object
17:         doprescrito(0) = op1
18:         doprescrito(1) = op2
19:         doprescrito(2) = op3
20:         doprescrito(3) = op4

21:         Dim novoarray(44) As Object
22:         novoarray(11) = cruzamento11
23:         novoarray(12) = cruzamento12
24:         novoarray(13) = cruzamento13
25:         novoarray(14) = cruzamento14
26:         novoarray(21) = cruzamento21
27:         novoarray(22) = cruzamento22
28:         novoarray(23) = cruzamento23
29:         novoarray(24) = cruzamento24
30:         novoarray(31) = cruzamento31
31:         novoarray(32) = cruzamento32
32:         novoarray(33) = cruzamento33
33:         novoarray(34) = cruzamento34
34:         novoarray(41) = cruzamento41
35:         novoarray(42) = cruzamento42
36:         novoarray(43) = cruzamento43
37:         novoarray(44) = cruzamento44

38:         Dim arrayc(3) As Object
            'webavi
            'arrayc(0) = C_PVP_1
            'arrayc(1) = C_PVP_2
            'arrayc(2) = C_PVP_3
            'arrayc(3) = C_PVP_4
39:         Dim dciAnosP(3) As Object
40:         dciAnosP(0) = quantosDCIa1nosP
41:         dciAnosP(1) = quantosDCIa2nosP
42:         dciAnosP(2) = quantosDCIa3nosP
43:         dciAnosP(3) = quantosDCIa4nosP

44:         Dim linha As Integer = 0
45:         For toshow As Integer = 0 To A - 1
46:             Dim index As String = (show(toshow).qualP) & (toshow + 1)

48:             If Not IsNothing(novoarray(index)) Then

49:                 If novoarray(index).cnpem = True Then

50:                     For secnpem As Integer = 1 To 5

51:                         If show(toshow).ranking.ToString.Substring(secnpem, 1) = 0 Or show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(secnpem, 1) = 2 Then 'só vai mudar se estiver a zero ou qty 1 ou 2
52:
                                If secnpem = 4 Then

53:                                  'marca não depende do CNPEM
55:                             ElseIf secnpem = 5 Then 'substitui o ranking da qty por certo no caso do cnpem ser o mesmo (para qty 0, 1 ou2)

56:                                 show(toshow).xqty = show(toshow).ranking.ToString.Substring(secnpem, 1)
57:                                 Mid(show(toshow).ranking, secnpem + 1, 1) = "3"
58:                             Else 'secnpem é 1, 2 ou 3 : dci, forma ou dose
59:                             'substitui o ranking da forma e dose por certo no caso do cnpem ser o mesmo

61:                                 Mid(show(toshow).ranking, secnpem + 1, 1) = "1"
62:                             End If
63:                         End If

                            'meter o xqty em vez do qty caso (5) <>3
64:                     Next
65:                 End If
66:             End If
341:
342:            Dim tempqty As Single = 0
343:
344:        'MsgBox(show(toshow).ranking)
345:            With reresult(toshow) 'richtextbox

346:                Dim dizera As String = ""
347:                Dim dizerp As String = ""
348:            'MsgBox(doprescrito(show(toshow).qualP - 1).code)
9999:               If show(toshow).qualP > 0 Then

                        If Not IsNothing(doprescrito(show(toshow).qualP - 1).nome) Then

349:                        If doprescrito(show(toshow).qualP - 1).code > 50000000 Then 'prescrito por cnpem

350:                            dizerp = UCase(doprescrito(show(toshow).qualP - 1).dci.substring(0, 1)) & doprescrito(show(toshow).qualP - 1).dci.substring(1, (Len(doprescrito(show(toshow).qualP - 1).dci) - 1))
351:                        Else

352:                            dizerp = UCase(doprescrito(show(toshow).qualP - 1).nome.substring(0, 1)) & doprescrito(show(toshow).qualP - 1).nome.substring(1, (Len(doprescrito(show(toshow).qualP - 1).nome) - 1))
353:                        End If
                        Else

                            dizerp = "desconhecido"
                            'GRAVARDESCONHECIDO
                        End If


354:                End If

355:                If show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(5, 1) = 2 Then

356:                    tempqty = show(toshow).ranking.ToString.Substring(5, 1)
357:                    Mid(show(toshow).ranking, 6, 1) = "0"
358:                ElseIf show(toshow).ranking.ToString.Substring(5, 1) = 1 Or show(toshow).ranking.ToString.Substring(5, 1) = 3 Then
359:                    tempqty = show(toshow).ranking.ToString.Substring(5, 1)
3561:               End If

3562:               If show(toshow).ranking = "0000000000" Then

                        If Not IsNothing(show(toshow).CNPEM) Then
412:                        If show(toshow).CNPEM > 0 Then
4121:                           For ixp As Integer = 0 To P - 1
4122:                               If doprescrito(ixp).cnpem = doaviado(toshow).CNPEM Then
4123:                                   .msg = doaviado(toshow).nome & " repetido"
9998:                                   .alinea = "k"
4124:                                   Exit For
4125:                               End If
4126:                               .msg = "z" & toshow + 1 & " = -> " & doaviado(toshow).nome & " não prescrito"
9997:                               .alinea = "z"
4127:                           Next
4128:                       Else
4129:                           For iixp As Integer = 0 To P - 1
4130:                               If doprescrito(iixp).dci = doaviado(toshow).dci Then
4131:                                   .msg = doaviado(toshow).nome & " repetido"
9996:                                   .alinea = "k"
4132:                                   Exit For
4133:                               End If
4134:                               .msg = "z" & toshow + 1 & " = -> " & doaviado(toshow).nome & " não prescrito"
9995:                               .alinea = "z"
4135:                           Next
4136:                       End If
                        Else
                            For pxi As Integer = 0 To P - 1
                                If doprescrito(pxi).code = doaviado(toshow).code Then
                                    .msg = "ok"
9980:                               .alinea = "a"
                                Else
                                    .msg = "desconhecido"
9970:                               .alinea = "9"
                                End If

                            Next
                        End If

4144:                   show(toshow).erro = True
415:                    verificador = "red"
                        'webavi final
                        '.selectionstart = 0
                        '.selectionlength = Len(reresult(toshow).msg)
                        .selectionbackcolor = verificador
4151:               Else
416:                    For posrank As Integer = 0 To 5


417:                        If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 AndAlso show(toshow).erro = False Then 'cnpem<>, dci<>, formaCNPEM<>, doseCNPEM<>, marca<>, qtyCNPEM<>
418:                         'MsgBox(show(toshow).ranking)
431:                            Select Case posrank  'ir concatenado += na richtextbox
                                    Case Is = 0
433:                                'cnpem diferente mas como não diz o porquê não mostra nada
434:                                Case Is = 1 'dci <>
435:                                    .msg += show(toshow).qualP & "y" & toshow + 1 & " = " & dizerp & "(" & doprescrito(show(toshow).qualP - 1).dci & ") -> " & doaviado(toshow).nome & "(" & doaviado(toshow).dci & ")"
436:                                    .alinea = "y"
4361:                                   show(toshow).erro = True
437:                                    verificador = "red"
438:                                Case Is = 2 'forma <>
439:                                    .msg += show(toshow).qualP & "f" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).FORMA & " -> " & doaviado(toshow).FORMA
440:                                    .alinea = "f"
441:                                    show(toshow).erro = True
4411:                                   verificador = "red"
442:                                Case Is = 3 'dose <>
443:                                    .msg += show(toshow).qualP & "g" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).dose & " -> " & doaviado(toshow).dose
444:                                    .alinea = "g"
4441:                                   show(toshow).erro = True
445:                                    verificador = "red"
446:                                Case Is = 4 'marca <>
447:                                    If Not doprescrito(show(toshow).qualP - 1).code > 50000000 Then 'para não mostrar marca dif quando prescrito por cnpem

448:
4481:                                       If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome) Or (doprescrito(show(toshow).qualP - 1).forma <> doaviado(toshow).forma) Then  'para não dar erro quando mesmo medicamento fica com novo código  mas para acusar creme -> pomada Then
4482:                                           If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome) Then
4483:                                               .msg += show(toshow).qualP & "s" & toshow + 1 & " = " & doprescrito(show(toshow).qualP - 1).nome & " -> " & doaviado(toshow).nome
449:                                                .alinea = "s"
4491:                                               show(toshow).erro = True
4492:                                               verificador = "red"
4493:                                           End If
4494:                                           If (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome) Then 'é forma diferente porque não há genéricos do prescrito
4495:                                               .msg += show(toshow).qualP & "f" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).FORMA & " -> " & doaviado(toshow).FORMA & " (porque não há genéricos do prescrito)"
4401:                                               .alinea = "f"
44011:                                              show(toshow).erro = True
4412:                                               verificador = "red"
44121:                                          End If

44122:                                      End If
44123:                                  Else

450:                                    End If
451:                                'show(toshow).erro = True 'como costuma estar mal mais vale não impedir de mostrar qtys
452:                                Case Is = 5
453:
454:                                    Select Case tempqty 'ver se h) ou L
                                            Case Is = 0 'h)
456:                                            .msg += show(toshow).qualP & "h" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " "
457:                                            .alinea = "h"
4571:                                           show(toshow).erro = True
4572:                                           verificador = "red"
458:                                        Case Is = 1 'x<y<150%
459:                                            .msg += show(toshow).qualP & "h" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " "
4591:                                           .alinea = "h"
4592:                                           show(toshow).erro = True
4593:                                           verificador = "red"
4601:                                       Case Is = 2 'L
4602:                                           .msg += show(toshow).qualP & "L" & toshow + 1 & " = " & dizerp & " " & doprescrito(show(toshow).qualP - 1).qty & " -> " & doaviado(toshow).qty & " "
4603:                                           .alinea = "L"
46031:                                          show(toshow).erro = True
46032:                                          If Not verificador = "red" Then
46033:                                              verificador = "orange"
46034:                                          End If
4604:                                       Case Else
4605:                                   End Select
4606:                               Case Else
4607:                           End Select
                                'webavi final
4608:                           '.selectionstart = 0
4609:                           '.selectionlength = Len(reresult(toshow).msg)
4610:                           .selectionbackcolor = verificador
4611:                       Else

4612:                       End If

4613:                   Next
461:                End If

462:                If show(toshow).erro = False Then   'a seguir acrescentei 0 a 17/12/2013
67:                     If show(toshow).ranking.ToString.Substring(0, 1) = 1 Or show(toshow).ranking.ToString.Substring(0, 1) = 0 Then 'prescrito por code com CNPEM dif e não encontrou erros
68:                         verificador = "red"
69:                         show(toshow).erro = True
70:                         .msg += dizerp & " " & doprescrito(show(toshow).qualP - 1).cnpem & " " & " CNPEM " & " " & doaviado(toshow).cnpem
71:                         .alinea = "w"
72:                'webavi final
464:                        '.selectionstart = 0
465:                        '.selectionlength = 26 + Len(dizerp)
73:                         .selectionbackcolor = verificador
74:                     Else
463:                        .msg += " OK " 'Environment.NewLine pode ser usado em vez do vbNewLine
75:                         .alinea = "a"
                            'webavi final
                            '.selectionstart = 0
                            '.selectionlength = 4
466:                        .selectionbackcolor = "green"
467:                    'verificador = Color.Green
76:                     End If
468:                End If
469:
470:
77:                 .cor = verificador
78:
79:             End With


                'verificar de aviados mais de dois iguais - só acusa se támbém mais de dois iguais prescritos (para não acusar desdobramentos)
                'webavi 
80:             If detectarmaisdedoisaviados() = True Then
81:                 verificador = "red"
82:                 If Not reresult(toshow).excepBoxText.Contains("mais de 2 iguais") Then
83:                     reresult(toshow).excepBoxText += "mais de 2 iguais"
84:                 End If
85:             End If


                If IsNothing(doprescrito(show(toshow).qualP - 1).nome) Then

                    For pppxi As Integer = 0 To P - 1
                        With reresult(toshow)
                            If doprescrito(pppxi).code = doaviado(toshow).code Then
                                .cor = "green"
                                .msg = "ok"

9960:                           .alinea = "a"

                                .top5 = False

                                .port = False

                                .excep = False


                                arrayoa(pppxi).dci = ""
                                arrayoa(pppxi).nome = ""
                                arrayoa(pppxi).forma = ""
                                arrayoa(pppxi).dose = ""

                                arrayoa(pppxi).qty = ""

                                arrayoa(pppxi).comp = 0

                                arrayoa(pppxi).gh = 0

                                arrayoa(pppxi).pvp = 0
                                arrayoa(pppxi).pr = 0
                                arrayoa(pppxi).gen = False
                                arrayoa(pppxi).lab = ""

                                arrayoa(pppxi).top5 = 0
                                arrayoa(pppxi).pvpmenos1 = 0
                                arrayoa(pppxi).pvpmenos2 = 0
                                arrayoa(pppxi).CNPEM = 0

                                arrayoa(pppxi).dci_obr = False
                                arrayoa(pppxi).pvpmenos3 = 0
                                arrayoa(pppxi).pvpmenos4 = 0
                                arrayoa(pppxi).pvpmenos5 = 0

                                arrayoa(pppxi).d4250 = False
                                arrayoa(pppxi).d1234 = False
                                arrayoa(pppxi).d21094 = False
                                arrayoa(pppxi).d10279 = False
                                arrayoa(pppxi).d10280 = False
                                arrayoa(pppxi).d10910 = False
                                arrayoa(pppxi).d14123 = False
                                arrayoa(pppxi).lei6 = False




                            Else
                                .cor = "red"
                                .msg = "desconhecido"
9950:                           .alinea = "9"
                                verificador = "red"
                                show(toshow).erro = True

                                .top5 = False
                                .port = False
                                .excep = False

                                arrayoa(pppxi).dci = ""
                                arrayoa(pppxi).nome = ""
                                arrayoa(pppxi).forma = ""
                                arrayoa(pppxi).dose = ""
                                arrayoa(pppxi).qty = ""
                                arrayoa(pppxi).comp = 0
                                arrayoa(pppxi).gh = 0
                                arrayoa(pppxi).pvp = 0
                                arrayoa(pppxi).pr = 0
                                arrayoa(pppxi).gen = False
                                arrayoa(pppxi).lab = ""
                                arrayoa(pppxi).top5 = 0
                                arrayoa(pppxi).pvpmenos1 = 0
                                arrayoa(pppxi).pvpmenos2 = 0
                                arrayoa(pppxi).CNPEM = 0
                                arrayoa(pppxi).dci_obr = False
                                arrayoa(pppxi).pvpmenos3 = 0
                                arrayoa(pppxi).pvpmenos4 = 0
                                arrayoa(pppxi).pvpmenos5 = 0
                                arrayoa(pppxi).d4250 = False
                                arrayoa(pppxi).d1234 = False
                                arrayoa(pppxi).d21094 = False
                                arrayoa(pppxi).d10279 = False
                                arrayoa(pppxi).d10280 = False
                                arrayoa(pppxi).d10910 = False
                                arrayoa(pppxi).d14123 = False
                                arrayoa(pppxi).lei6 = False




                            End If
                        End With
                    Next
                End If

86:             With reresult(toshow)
87:                 If show(toshow).ranking.ToString.Substring(8, 1) = 1 Or 2 Then 'mb ou mp  
471:                    If Not (show(toshow).ranking.ToString.Substring(0, 1) = "0" Or "1") Then 'para não mostrar mais barato quando prescrito por cnpem
472:                        If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome And show(toshow).ranking.ToString.Substring(5, 1) < 3) Then 'para não acusar excep quando mesmo medicamento com L ou H
88:                             If verificador = "green" Then
89:                                 verificador = "yellow"

90:                             End If

91:                             .excepBoxtext += show(toshow).qualP & "t " & toshow + 1 '& " "

                                'webavi final
473:                            '.excepBoxSelectionStart = Len(reresult(toshow).msg.ToString)
474:                            '.excepBoxSelectionLength = 5
475:                            .excepBoxSelectionBackColor = verificador
94:                         End If
95:                     End If
96:                 End If


                    ' For posrank As Integer = 6 To 9
                    'If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 Then 'a), b), c) ou (>)
                    '             If verificador = Color.Green Then
                    'verificador = Color.Yellow
4901:           'End If
4902:

4903:               If show(toshow).erro = False Then
4904:                   If show(toshow).ranking.ToString.Substring(6, 1) = 0 Or show(toshow).ranking.ToString.Substring(7, 1) = 0 Or show(toshow).ranking.ToString.Substring(8, 1) = 0 Then
4905:                       If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome And show(toshow).ranking.ToString.Substring(5, 1) < 3) Then 'para não acusar excep quando mesmo medicamento com L ou H
4906:                       'linha = linha + 1

99:                             If Not (doprescrito(show(toshow).qualP - 1).nome = doaviado(toshow).nome) Then  'para não dar erro quando mesmo medicamento fica com novo código  Then


4907:                               .ExcepBoxText += show(toshow).qualP & " " & doaviado(toshow).dci & " " & toshow + 1
100:                                With reresult(toshow)
4908:                                   .msg += show(toshow).qualP & " c) " & toshow + 1

101:                                    totalresult.excep = True
102:                                    .excep = True
                                        'webavi final
4909:                               '.selectionstart = reresult(toshow).msg.IndexOf(reresult(toshow).Lines(1))
4910:                               '.selectionlength = 6
4911:                                   .selectionbackcolor = "yellow"
                                    End With
4912: 'webavi
                                    '  arrayc(toshow).Text = doprescrito(show(toshow).qualP - 1).PVPexcepc
4913:                               If verificador = "green" Then
4914:                                   verificador = "yellow"
4915:                               End If 'fim este
                                End If 'fim nome dif
4916:                       End If 'fim com L ou H

4917:                   End If 'fim substring
4918:               End If 'fim erro
4919:           ' .Lines = New String() {linha - 1}
150:             ' .SelectionStart = 0
                    '.ExcepBoxSelectionLength = Len(doaviado(toshow).dci) + 4
                    'If posrank = 9 Then
                    '    .ExcepBoxSelectionLength = 8
                    'End If
                    '.ExcepBoxSelectionBackColor = "yellow"
151:             'End If
152:             ' Next
153:
154:            End With




155:            With reresult(toshow)
156:                For posrank As Integer = 6 To 9
157:                    If show(toshow).erro = False Then
158:                        If show(toshow).ranking.ToString.Substring(posrank, 1) = 0 Then 'a), b), c) ou (>)
                                'os das excepções estão na excepbox
160:                            Select Case posrank
                                    Case Is = 9
162:                                    .msg += "<br/> >top5 " & toshow + 1 '& " "
163:                                    .cor = "yellow"
164:                                    .top5 = True
165:                                    totalresult.top5 = True

167:                            End Select
                                'webavi final
                                ' .top5BoxSelectionStart = 0
                                '.top5BoxSelectionLength = 7
                                '.top5BoxSelectionBackColor = verificador
170:                        End If
171:                    End If
172:                Next
173:            End With



174:        Next
            'WebAvi

301:
303:        For iiii As Integer = 0 To A - 1
303303:
304:            If doaviado(iiii).d10910 = True Or doaviado(iiii).d4250 = True Or doaviado(iiii).d1234 = True Or doaviado(iiii).d10279 = True Or doaviado(iiii).d10280 = True Or doaviado(iiii).d21094 = True Or doaviado(iiii).d14123 = True Or doaviado(iiii).lei6 = True Then 'vê se aviado algum código ao qual possam ser atribuidas portarias
305:                reresult(iiii).port = True
306:                totalresult.port = True
                    '175:                If verificador = "green" Then   'comentei para poder não ficar amarelo em lotes que não 1 e 48
                    '176:                    verificador = "yellow"
                    '177:                End If

                    '178:                With reresult(iiii)
                    '179:                    If doaviado(iiii).d10910 = True Then
                    '180:                        .msg += "<br/> Despacho n.º 10910/2011, de 22/04 ?"
                    '181:                    End If
                    '182:                    If doaviado(iiii).d4250 = True Then
                    '183:                        .msg += "<br/> Despacho n.º 13020/2011, de 20/09 ?"
                    '184:                    End If
                    '185:                    If doaviado(iiii).d21094 = True Then
                    '186:                        .msg += "<br/> Despacho n.º 21094/99, de 14/09 ?"
                    '187:                    End If
                    '188:                    If doaviado(iiii).lei6 = True Then
                    '189:                        .msg += "<br/> Lei n.º 6/2010, de 07/05 ?"
                    '190:                    End If
                    '191:                    If doaviado(iiii).d10279 = True Then
                    '192:                        .msg += "<br/> Despacho n.º 10279/2008, de 11/03 ?"
                    '193:                        If doaviado(iiii).d10280 = True Then
                    '194:                            .msg += " e Despacho n.º 10280/2008, de 11/03 ?"
                    '195:           End If
                    '196:                    ElseIf doaviado(iiii).d10280 = True Then
                    '197:                        .msg += " <br/> Despacho n.º 10280/2008, de 11/03 ?"
                    '198:                    End If

                    '199:                    If doaviado(iiii).d14123 = True Then
                    '200:                        .msg += "<br/> Despacho n.º 14123/2009(2ª série), de 12/06 ?"
                    '201:                        If doaviado(iiii).d1234 = True Then
                    '202:                            .msg += " e Despacho n.º 1234/2006, de 29/12 ?"
                    '203:                        End If
                    '204:                    ElseIf doaviado(iiii).d1234 = True Then
                    '205:                        .msg += " <br/> Despacho n.º 1234/2006, de 29/12 ?"
                    '206:                    End If

                    '217:                End With
307:                End If

308:        Next
309:
310:        totalresult.cor = verificador




            'falta ARRANJAR SíTIO PARA ISTO
146:    'If aceitarduplicados = True Then
1350:    '   'só usado na versão lite (aceita repetições, iguais e desdobramentos de qualquer tamanho)
365:    '  For AAA As Integer = 1 To A
369:    '    If AAA <> iterAAA Then
            'If (arrayoa(AAA - 1).cnpem > 0 And (arrayoa(AAA - 1).cnpem = arrayoa(iterAAA - 1).cnpem)) Or (arrayoa(AAA - 1).code = arrayoa(iterAAA - 1).code) Then
370:    '    If iterAAA > AAA Then 'para não fazer duas vezes e desfazer o que estava feito
            'cruzamentocerto(iterAAA - 1).ranking = cruzamentocerto(AAA - 1).ranking
            'End If
372:    'End If
373:    'End If
374:    'Next
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


311:        For zz As Integer = 0 To A - 1
312:            If reresult(zz).alinea = "" Then
313:                reresult(zz).alinea = reresult(zz).msg.substring(1, 1)
314:            End If
315:  Select Case reresult(zz).alinea
                        Case "z"
316:                    If Not (totalresult.verificador >= 5) Then
317:                        totalresult.alinea = "z"
318:                        totalresult.verificador = 4
319:                    End If
320:                Case "y"
321:                    totalresult.alinea = "y"
322:                    totalresult.verificador = 9
323:                Case "w"
324:                    If Not (totalresult.verificador >= 7) Then
325:                        totalresult.alinea = "w"
326:                        totalresult.verificador = 6
327:                    End If
328:                Case "s"
329:                    If Not (totalresult.verificador >= 6) Then
330:                        totalresult.alinea = "s"
331:                        totalresult.verificador = 5
332:                    End If
333:                Case "k"
334:                    If Not (totalresult.verificador >= 5) Then
335:                        totalresult.alinea = "k"
336:                        totalresult.verificador = 4
337:                    End If
338:                Case "f"
339:                    If Not (totalresult.verificador >= 8) Then
340:                        totalresult.alinea = "f"
361:                        totalresult.verificador = 7
362:                    End If
363:                Case "g"
364:                    If Not (totalresult.verificador >= 8) Then
36512:                      totalresult.alinea = "g"
366:                        totalresult.verificador = 8
367:                    End If
368:                Case "h"
379:                    If Not (totalresult.verificador >= 4) Then
380:                        totalresult.alinea = "h"
381:                        totalresult.verificador = 3
382:                    End If
383:                Case "L"
384:                    If Not (totalresult.verificador >= 3) Then
385:                        totalresult.alinea = "L"
386:                        totalresult.verificador = 2
387:                    End If
388:                Case "a"
389:                    If totalresult.alinea = "" Then
390:                        totalresult.alinea = "a"
391:                        totalresult.verificador = 0
392:                    End If
393:            End Select

394:        Next

395:        resultadomostrado = True

396:        For ww As Integer = 0 To A - 1
397:            totalresult.msg += reresult(ww).msg
398:        Next



            Exit Function
MOSTRARERRO:
            MsgBox("F08" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function

        <HttpPost>
        Function mostrar(receita As Object, rlinha1 As Object, rlinha2 As Object, rlinha3 As Object, rlinha4 As Object, medsp1 As Object, medsp2 As Object, medsp3 As Object, medsp4 As Object, medsa1 As Object, medsa2 As Object, medsa3 As Object, medsa4 As Object) As ActionResult
            On Error GoTo MOSTRARERRO
            ' ViewBag.cor = "red"
            ' ViewBag.alinea = "s"

            'arrayfinal(0) = receita
            'arrayfinal(1) = rlinha1


            'Return (cor, alinea)'swapnil commented
            'swapnil added 

            Dim lst() As Object = {receita, rlinha1, rlinha2, rlinha3, rlinha4, medsp1, medsp2, medsp3, medsp4, medsa1, medsa2, medsa3, medsa4}
            'mostrar(array1, array) 'swapnil commented
            'swapnil added 
            Return Me.Json(lst.ToArray, JsonRequestBehavior.AllowGet)




            Exit Function
MOSTRARERRO:
            MsgBox("F17" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function

        Function SUBROTINA2(ByVal ITERppp As Single, ByVal ITERaaa As Single, ByVal ITERadorprior As Single, ByVal iteradorQual1 As SByte, ByVal iteradorQual2 As SByte, ByVal arraycroces As Array, ByVal cruzamentocerto As Array, ByVal quantosdciAnosP As Single, ByVal quantosdciAnosA As Single)
1:          On Error GoTo MOSTRARERRO
            cruzamentocerto(0) = resultado1
            cruzamentocerto(1) = resultado2
            cruzamentocerto(2) = resultado3
            cruzamentocerto(3) = resultado4
2:          Dim ordenados(3) As Object
3:          ordenados(0) = ordenado1
4:          ordenados(1) = ordenado2
5:          ordenados(2) = ordenado3
6:          ordenados(3) = ordenado4


            Dim arraycruza(44) As Object
            arraycruza(11) = cruzamento11
7:          arraycruza(12) = cruzamento12
8:          arraycruza(13) = cruzamento13
9:          arraycruza(14) = cruzamento14
10:         arraycruza(21) = cruzamento21
11:         arraycruza(22) = cruzamento22
12:         arraycruza(23) = cruzamento23
13:         arraycruza(24) = cruzamento24
14:         arraycruza(31) = cruzamento31
15:         arraycruza(32) = cruzamento32
16:         arraycruza(33) = cruzamento33
17:         arraycruza(34) = cruzamento34
18:         arraycruza(41) = cruzamento41
19:         arraycruza(42) = cruzamento42
20:         arraycruza(43) = cruzamento43
21:         arraycruza(44) = cruzamento44


93:         ITERppp = 1
94:         Do While ITERppp <= P
361:
363:            ITERadorprior = Convert.ToSingle((ITERppp) & (ITERaaa))
37:             If arraycroces(ITERadorprior).anulado = False Then
38:                 If arraycroces(ITERadorprior).ranking.ToString.Substring(1) >= cruzamentocerto(ITERaaa - 1).ranking.ToString.Substring(1) Then
39:                     If Not (quantosdciAnosP - quantosdciAnosA) < 0 Then
390:                        cruzamentocerto(ITERaaa - 1).qualP = ITERppp
391:                        cruzamentocerto(ITERaaa - 1).ranking = arraycroces(ITERadorprior).ranking
393:                        iteradorQual1 = Convert.ToSByte(cruzamentocerto(ITERaaa - 1).qualP & ITERaaa)
394:                        iteradorQual2 = iteradorQual1
395:                    Else         'mais aviados que prescritos 
396:                    'ordeno por ranking
                            ReDim ordenados(ITERaaa - 1)(P - 1)
397:                        For o As Integer = 0 To P - 1 'quantosdciAnosA - 1
                                Dim concata1 = (o + 1) & A
398:                            concatenado = arraycruza(concata1).ranking & concata1
399:                            ordenados(ITERaaa - 1)(o) = concatenado
4001:                       '  MsgBox((ITERaaa - 1) & " " & o & " " & ordenados(ITERaaa - 1)(o))
4002:                       Next
4003:
401:                    'nos dois a seguir era A em vez de P e mudei a 17/12/2013
402:                        Array.Sort(ordenados(ITERaaa - 1), 0, P) 'ordenar (é por ordem crescente)
403:                        Array.Reverse(ordenados(ITERaaa - 1), 0, P) 'fica decrescente

408:                        For q As Integer = 1 To quantosdciAnosP 'quantos A nos P diz quantos de maior ranking
409:                        'os ordenadoss de maior ranking passam como estão
                                'aaaa0:                          MsgBox("iteraaa-1: " & ITERaaa - 1)
                                'aaaa1:                          MsgBox("q-1: " & q - 1)
                                'aaaa2:                          MsgBox("iterppp: " & ITERppp)
                                'aaaa3:                          MsgBox("ordenados(ITERaaa - 1)(q - 1)).ToString: " & (ordenados(ITERaaa - 1)(q - 1)).ToString)
                                'aaaa4:                          MsgBox("ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)): " & (ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1))
                                'aaaa5:                          MsgBox("(ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1): " & (ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1))
                                'aaaa6:                          MsgBox("cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp: " & cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp)
410:                            cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp = ITERppp
411:                        'MsgBox(cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).RANKING)
415:                            cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).RANKING = arraycroces(ITERadorprior).ranking
416:                            iteradorQual1 = ordenados(ITERaaa - 1)(q - 1).ToString.Substring(10, 2) 'cruzamentocerto((ordenados(ITERaaa - 1)(q - 1)).ToString.Substring(10, 1)).qualp & ... 'ERRO 13 = o ordenados não dá número, falta-me o indice
417:                            iteradorQual2 = iteradorQual1
419:                        Next

420:                    End If
40:                 Else
41:                 'fica como estava
42:                 End If
43:             Else
44:             'fica como estava
45:             End If
46:
47:
48:             ITERppp = ITERppp + 1
49:         Loop 'final do while iterPPP

50:         anular(iteradorQual1, iteradorQual2, ITERaaa, arraycroces, cruzamentocerto)


54:         Exit Function
MOSTRARERRO:
            MsgBox("F09" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function



        Function anular(ByVal iteradorqual1 As SByte, iteradorqual2 As SByte, ByVal iterAAA As Single, ByVal arraycroces As Object, ByVal cruzamentocerto As Object)
            On Error GoTo MOSTRARERRO

6:          arraycroces(11) = cruzamento11
7:          arraycroces(12) = cruzamento12
8:          arraycroces(13) = cruzamento13
9:          arraycroces(14) = cruzamento14
10:         arraycroces(21) = cruzamento21
11:         arraycroces(22) = cruzamento22
12:         arraycroces(23) = cruzamento23
13:         arraycroces(24) = cruzamento24
14:         arraycroces(31) = cruzamento31
15:         arraycroces(32) = cruzamento32
16:         arraycroces(33) = cruzamento33
17:         arraycroces(34) = cruzamento34
18:         arraycroces(41) = cruzamento41
19:         arraycroces(42) = cruzamento42
20:         arraycroces(43) = cruzamento43
21:         arraycroces(44) = cruzamento44

            iteradorqual1 = iteradorqual1 - 30

50:         Do While iteradorqual1 <= P * 10 + iterAAA
                If iteradorqual1 > 10 Then
51:                 arraycroces(iteradorqual1).anulado = True
52:             End If
                iteradorqual1 = iteradorqual1 + 10
53:         Loop
54:
55:         Do While iteradorqual2 <= cruzamentocerto(iterAAA - 1).qualP * 10 + A
                If aceitarduplicados = False Then
56:                 arraycroces(iteradorqual2).anulado = True
57:                 iteradorqual2 = iteradorqual2 + 1
                End If
58:         Loop
            Exit Function
MOSTRARERRO:
            MsgBox("F10" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function



        Function makeranking2013()
1:          On Error GoTo MOSTRARERRO


2:          Dim iterPP As Single = 1
3:          Dim iterAA As Single = 1
            Dim arraycruzamentos(44) As Object
5:          arraycruzamentos(11) = cruzamento11
6:          arraycruzamentos(12) = cruzamento12
7:          arraycruzamentos(13) = cruzamento13
8:          arraycruzamentos(14) = cruzamento14
9:          arraycruzamentos(21) = cruzamento21
10:         arraycruzamentos(22) = cruzamento22
11:         arraycruzamentos(23) = cruzamento23
12:         arraycruzamentos(24) = cruzamento24
13:         arraycruzamentos(31) = cruzamento31
14:         arraycruzamentos(32) = cruzamento32
15:         arraycruzamentos(33) = cruzamento33
16:         arraycruzamentos(34) = cruzamento34
17:         arraycruzamentos(41) = cruzamento41
18:         arraycruzamentos(42) = cruzamento42
19:         arraycruzamentos(43) = cruzamento43
20:         arraycruzamentos(44) = cruzamento44
            Dim arrayp(3) As Object
            arrayp(0) = op1
            arrayp(1) = op2
            arrayp(2) = op3
            arrayp(3) = op4
            Dim ARRAYdciP(3) As String
            Dim iteradorMR As Single = 11

21:         Do While iterAA <= A
                iterPP = 1
22:             Do While iterPP <= P
                    iteradorMR = Convert.ToSingle((iterPP) & (iterAA))
23:             'ranking do code c*********
24:                 If arraycruzamentos(iteradorMR).code = True Then 'se mesmo código
25:                     arraycruzamentos(iteradorMR).ranking = "411123"

26:                     GoTo OK
27:                 ElseIf arrayp(iterPP - 1).porCNPEM = False Then 'se não prescrito por cnpem
                        'será como se xcnpem =1 ou 3
272:                    If arrayp(iterPP - 1).cnpem = 0 Then 'prescrito por code diferente e sem CNPEM
273:                        arraycruzamentos(iteradorMR).ranking = "2" 'comparar à antiga, pode estar certo ou errado
274:                    Else 'prescrito por code diferente com CNPEM
275:                        If arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por code dif com mesmo cnpem (pode estar certo ou S)
276:                        'arraycruzamentos(iteradorMR).ranking = "2111" & arraycruzamentos(iteradorMR).marcadifsemhavergens & "3"
                                'GoTo Ok
                                '[se deixasse como está em cima, no caso da marcadif ser 1, o showresult teria de ser alterado para detectar como erro]
                                If arraycruzamentos(iteradorMR).marcadifsemhavergens = 2 Then
277:                                arraycruzamentos(iteradorMR).ranking = "211123"
                                    GoTo OK
279:                            Else
280:                                arraycruzamentos(iteradorMR).ranking = "211103" 'marcadif (0 ou 1)
                                    GoTo OK
281:                            End If
282:                        Else
28:                             arraycruzamentos(iteradorMR).ranking = "1"  'por code dif com cnpem dif - está errado -  vai seguir o restante
29:                         End If
291:                    End If
292:                ElseIf arraycruzamentos(iteradorMR).cnpem = True Then 'se prescrito por cnpem e igual 'está diferente no não lite
30:                     arraycruzamentos(iteradorMR).ranking = "311123"
31:                     GoTo OK
32:                 Else 'prescrito por cnpem e diferente (xcnpem = 2 ou 0)
33:                     arraycruzamentos(iteradorMR).ranking = "0" 'prescrito por cnpem e diferente

34:                 End If
35:
36:             'ranking do dci *d********
37:                 If arraycruzamentos(iteradorMR).dci = True Then
38:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
39:                 Else
40:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
41:                 End If
42:
43:             'ranking da forma **f*******
44:                 If arraycruzamentos(iteradorMR).forma = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'era diferente no não lite

45:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
46:                 Else
47:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
48:                 End If
49:
50:             'ranking da dose ***d******
51:                 If arraycruzamentos(iteradorMR).dose = True Or arraycruzamentos(iteradorMR).cnpem = True Then 'era diferente no não lite
52:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "1"
53:                 Else
54:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "0"
55:                 End If
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
70:                     arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).excepa & arraycruzamentos(iteradorMR).excepb & arraycruzamentos(iteradorMR).excepc
71:                 Else 'prescrito por CNPEM
                        arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & "555"
                    End If

72:             'ranking do top5 *********t
73:                 arraycruzamentos(iteradorMR).ranking = arraycruzamentos(iteradorMR).ranking & arraycruzamentos(iteradorMR).top5
74:
75:
76:                 iterPP = iterPP + 1
77:             Loop
78:             iterAA = iterAA + 1
79:         Loop
            novoprioridade2013()
80:         Exit Function
MOSTRARERRO:
            MsgBox("F11" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

        End Function



        Function comparador2013(ByVal cmprdrP As Object, ByVal cmprdrA As Object)
1:          On Error GoTo MOSTRARERRO




            Dim posicaop1(3) As Single
            Dim posicaop2(3) As Single
            Dim posicaoa1(3) As Single
            Dim posicaoa2(3) As Single
2:          Dim arrayA(3) As Object
3:          Dim arrayP(3) As Object
4:          Dim arraycruz(44) As Object '{11, 12, 13, 14, 21, 22, 23, 24, 31, 32, 33, 34, 41, 42, 43, 44}
6:          Dim iterP As Single = 1
7:          Dim iterA As Single = 1
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
9:          arrayA(0) = oa1
10:         arrayA(1) = oa2
11:         arrayA(2) = oa3
12:         arrayA(3) = oa4
13:         arrayP(0) = op1
14:         arrayP(1) = op2
15:         arrayP(2) = op3
16:         arrayP(3) = op4
17:         arraycruz(11) = cruzamento11
18:         arraycruz(12) = cruzamento12
19:         arraycruz(13) = cruzamento13
20:         arraycruz(14) = cruzamento14
21:         arraycruz(21) = cruzamento21
22:         arraycruz(22) = cruzamento22
23:         arraycruz(23) = cruzamento23
24:         arraycruz(24) = cruzamento24
25:         arraycruz(31) = cruzamento31
26:         arraycruz(32) = cruzamento32
27:         arraycruz(33) = cruzamento33
28:         arraycruz(34) = cruzamento34
29:         arraycruz(41) = cruzamento41
30:         arraycruz(42) = cruzamento42
31:         arraycruz(43) = cruzamento43
32:         arraycruz(44) = cruzamento44


            'MsgBox("03" & vbCr & "" & vbCr & "cmprdrP1: " & cmprdrP1 & "          cmprdrA1: " & cmprdrA1 & vbCr & "cmprdrP2: " & cmprdrP2 & "          cmprdrA2: " & cmprdrA2 & vbCr & "cmprdr3: " & cmprdrP3 & "          cmprdrA3: " & cmprdrA3 & vbCr & "cmprdrP4: " & cmprdrP4 & "          cmprdrA4: " & cmprdrA4 & " .")

            Dim iterador As Single = 11
33:         Do While iterA <= A
                iterP = 1
34:             Do While iterP <= P
                    iterador = Convert.ToSingle((iterP) & (iterA))
35:                 If arrayA(iterA - 1).code = arrayP(iterP - 1).code Then
36:                     arraycruz(iterador).code = True
37:                 Else
38:                     arraycruz(iterador).code = False
39:                 End If

40:                 If arrayA(iterA - 1).dci = arrayP(iterP - 1).dci Then
41:                     arraycruz(iterador).dci = True
42:                 Else
43:                     arraycruz(iterador).dci = False
44:                 End If

45:                 If arrayA(iterA - 1).nome = arrayP(iterP - 1).nome Then
46:                     arraycruz(iterador).nome = True
47:                 Else
48:                     arraycruz(iterador).nome = False
49:                 End If

50:                 If arrayA(iterA - 1).forma = arrayP(iterP - 1).forma Then
51:                     arraycruz(iterador).forma = True
52:                 Else
53:                     arraycruz(iterador).forma = False
54:                 End If

68:                 If arraycruz(iterador).forma = False Then 'caso particular das formas que podem ser equivalentes (não implica que a regra do CNPEM não acuse erro por causa da forma, por exemplo creme e pomada)
69:                     If via(arrayP(iterP - 1).forma) = via(arrayA(iterA - 1).forma) Then
70:                         arraycruz(iterador).forma = True
71:                     End If
72:                 End If

721:                If arrayA(iterA - 1).dose = arrayP(iterP - 1).dose Then
722:                    arraycruz(iterador).dose = True
723:                Else
724:                    arraycruz(iterador).dose = False
725:                End If

726:                If arrayA(iterA - 1).qty = arrayP(iterP - 1).qty Then
727:                    arraycruz(iterador).qty = True
728:                Else
729:                    arraycruz(iterador).qty = False
730:                End If

73:                 If IsNumeric(arrayP(iterP - 1).qty) AndAlso IsNumeric(arrayA(iterA - 1).qty) Then 'qty numerica em ambos
731:                    If arraycruz(iterador).qty = False Then 'se a quantidade não é igual vai ver se é menor, maior e maior que 50% (independentemente CNPEM)
74:                         If Convert.ToInt16(arrayP(iterP - 1).qty) > Convert.ToInt16(arrayA(iterA - 1).qty) Then 'se qty aviado < prescrito
75:                             arraycruz(iterador).xqty = 2
76:                         ElseIf Convert.ToInt16(arrayP(iterP - 1).qty) > 1.5 * Convert.ToInt16(arrayA(iterA - 1).qty) Then 'se qty aviado > 150% prescrito
77:                             arraycruz(iterador).xqty = 0
78:                         ElseIf Convert.ToInt16(arrayP(iterP - 1).qty) < 1.5 * Convert.ToInt16(arrayA(iterA - 1).qty) And Convert.ToInt16(arrayA(iterA - 1).qty) > Convert.ToInt16(arrayP(iterP - 1).qty) Then 'se qty prescrito < qty aviado < 150 % prescrito
79:                             arraycruz(iterador).xqty = 1
80:                         End If
81:                     Else
82:                         arraycruz(iterador).xqty = 3 'qty =     (independentemente CNPEM) 
83:                     End If

801:                Else 'qty não numerica em pelo menos um dos dois
                        If IsNumeric(arrayP(iterP - 1).qty) Or IsNumeric(arrayA(iterA - 1).qty) Then 'qty numerica num deles
                            'é erro de apresentacao.
8147:                       arraycruz(iterador).xqty = 0
                        Else 'nenhum é numerico
                            If Len(arrayP(iterP - 1).qty) >= 17 Then
802:                            posicaop1(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("-")
803:                            If arrayP(iterP - 1).qty.ToString.Contains("o") Then
804:                                posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("o")
805:                            ElseIf arrayP(iterP - 1).qty.ToString.Contains("m") Then
806:                                posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("m")
807:                            ElseIf arrayP(iterP - 1).qty.ToString.Contains("g") Then
808:                                posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("g")
809:                            ElseIf arrayP(iterP - 1).qty.ToString.Contains("l") Then
810:                                posicaop2(iterP - 1) = arrayP(iterP - 1).qty.ToString.IndexOf("l")
811:                            End If
812:                            antesp(iterP - 1) = arrayP(iterP - 1).qty.ToString.Substring(0, posicaop1(iterP - 1) - 12)
813:                            depoisp(iterP - 1) = arrayP(iterP - 1).qty.ToString.Substring(posicaop1(iterP - 1) + 2, posicaop2(iterP - 1) - posicaop1(iterP - 1) - 3)
814:
                            End If
                            If Len(arrayA(iterA - 1).qty) >= 17 Then
815:                            posicaoa1(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("-")
816:                            If arrayA(iterA - 1).qty.ToString.Contains("o") Then
817:                                posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("o")
818:                            ElseIf arrayA(iterA - 1).qty.ToString.Contains("m") Then
819:                                posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("m")
820:                            ElseIf arrayA(iterA - 1).qty.ToString.Contains("g") Then
821:                                posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("g")
822:                            ElseIf arrayA(iterA - 1).qty.ToString.Contains("l") Then
823:                                posicaoa2(iterA - 1) = arrayA(iterA - 1).qty.ToString.IndexOf("l")
824:                            End If

825:                            antesa(iterA - 1) = arrayA(iterA - 1).qty.ToString.Substring(0, posicaoa1(iterA - 1) - 12)
826:                            depoisa(iterA - 1) = arrayA(iterA - 1).qty.ToString.Substring(posicaoa1(iterA - 1) + 2, posicaoa2(iterA - 1) - posicaoa1(iterA - 1) - 3)
827:                        End If
828:                    'compara antesp(iterP - 1) com antesa(iterA - 1)
829:                        If antesp(iterP - 1) <> antesa(iterA - 1) Then 'se a quantidade não numerica não é igual vai ver se é menor, maior e maior que 50% (independentemente CNPEM)
830:                            If Convert.ToInt16(antesp(iterP - 1)) > Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado < prescrito
831:                                arraycruz(iterador).xqty = 2
832:                            ElseIf Convert.ToInt16(antesp(iterP - 1)) > 1.5 * Convert.ToInt16(antesa(iterA - 1)) Then 'se qty aviado > 150% prescrito
833:                                arraycruz(iterador).xqty = 0
834:                            ElseIf Convert.ToInt16(antesp(iterP - 1)) < 1.5 * Convert.ToInt16(antesa(iterA - 1)) And Convert.ToInt16(antesa(iterA - 1)) > Convert.ToInt16(antesp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
835:                                arraycruz(iterador).xqty = 1
                                Else
                                    arraycruz(iterador).xqty = 3
8136:                           End If
8137:                       Else 'se antes é igual compara depoisp(iterP - 1) com depoisa(iterA - 1)
8138:                           If Convert.ToInt16(depoisp(iterP - 1)) > Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado < prescrito
8139:                               arraycruz(iterador).xqty = 2
8140:                           ElseIf Convert.ToInt16(depoisp(iterP - 1)) > 1.5 * Convert.ToInt16(depoisa(iterA - 1)) Then 'se qty aviado > 150% prescrito
8141:                               arraycruz(iterador).xqty = 0
8142:                           ElseIf Convert.ToInt16(depoisp(iterP - 1)) < 1.5 * Convert.ToInt16(depoisa(iterA - 1)) And Convert.ToInt16(depoisa(iterA - 1)) > Convert.ToInt16(depoisp(iterP - 1)) Then 'se qty prescrito < qty aviado < 150 % prescrito
8143:                               arraycruz(iterador).xqty = 1
                                Else
                                    arraycruz(iterador).xqty = 3
8144:                           End If
8145:                       End If
8146:                   End If
8148:               End If



8231:               If arrayA(iterA - 1).comp = arrayP(iterP - 1).comp Then
8232:                   arraycruz(iterador).comp = True
8233:               Else
8234:                   arraycruz(iterador).comp = False
8235:               End If

836:                If arrayA(iterA - 1).GH = arrayP(iterP - 1).GH Then
837:                    arraycruz(iterador).GH = True
838:                Else
839:                    arraycruz(iterador).GH = False
840:                End If

841:                If arrayA(iterA - 1).pr = arrayP(iterP - 1).pr Then
842:                    arraycruz(iterador).pr = True
843:                Else
844:                    arraycruz(iterador).pr = False
845:                End If

846:                If arrayA(iterA - 1).gen = arrayP(iterP - 1).gen Then
847:                    arraycruz(iterador).gen = True
848:                Else
849:                    arraycruz(iterador).gen = False
850:                End If

851:                If arrayA(iterA - 1).lab = arrayP(iterP - 1).lab Then
852:                    arraycruz(iterador).lab = True
853:                Else
854:                    arraycruz(iterador).lab = False
855:                End If

856:
861:                If arrayA(iterA - 1).dci_obr = arrayP(iterP - 1).dci_obr Then
862:                    arraycruz(iterador).dci_obr = True
863:                Else
864:                    arraycruz(iterador).dci_obr = False
865:                End If

866:                If arrayA(iterA - 1).d4250 = arrayP(iterP - 1).d4250 Then
867:                    arraycruz(iterador).d4250 = True
868:                Else
869:                    arraycruz(iterador).d4250 = False
870:                End If

871:                If arrayA(iterA - 1).d1234 = arrayP(iterP - 1).d1234 Then
872:                    arraycruz(iterador).d1234 = True
873:                Else
874:                    arraycruz(iterador).d1234 = False
875:                End If

876:                If arrayA(iterA - 1).d21094 = arrayP(iterP - 1).d21094 Then
877:                    arraycruz(iterador).d21094 = True
878:                Else
879:                    arraycruz(iterador).d21094 = False
880:                End If

881:                If arrayA(iterA - 1).d10279 = arrayP(iterP - 1).d10279 Then
882:                    arraycruz(iterador).d10279 = True
883:                Else
884:                    arraycruz(iterador).d10279 = False
885:                End If

886:                If arrayA(iterA - 1).d10280 = arrayP(iterP - 1).d10280 Then
887:                    arraycruz(iterador).d10280 = True
888:                Else
889:                    arraycruz(iterador).d10280 = False
890:                End If

891:                If arrayA(iterA - 1).d10910 = arrayP(iterP - 1).d10910 Then
892:                    arraycruz(iterador).d10910 = True
893:                Else
894:                    arraycruz(iterador).d10910 = False
895:                End If

896:                If arrayA(iterA - 1).d14123 = arrayP(iterP - 1).d14123 Then
897:                    arraycruz(iterador).d14123 = True
898:                Else
899:                    arraycruz(iterador).d14123 = False
900:                End If

901:                If arrayA(iterA - 1).lei6 = arrayP(iterP - 1).lei6 Then
902:                    arraycruz(iterador).lei6 = True
903:                Else
904:                    arraycruz(iterador).lei6 = False
905:                End If

906:                If arrayA(iterA - 1).pvpmenos1 = arrayP(iterP - 1).pvpmenos1 Then
907:                    arraycruz(iterador).pvpmenos1 = True
908:                Else
909:                    arraycruz(iterador).pvpmenos1 = False
910:                End If

911:                If arrayA(iterA - 1).pvpmenos2 = arrayP(iterP - 1).pvpmenos2 Then
912:                    arraycruz(iterador).pvpmenos2 = True
913:                Else
914:                    arraycruz(iterador).pvpmenos2 = False
915:                End If

916:                If arrayA(iterA - 1).trocamarca = arrayP(iterP - 1).trocamarca Then
917:                    arraycruz(iterador).trocamarca = True
918:                Else
919:                    arraycruz(iterador).trocamarca = False
920:                End If
921:
922:                If arrayA(iterA - 1).cnpem = arrayP(iterP - 1).cnpem Then
923:                    arraycruz(iterador).cnpem = True
924:                Else
925:                    arraycruz(iterador).cnpem = False
926:                End If

927:                If arrayA(iterA - 1).nDCI = arrayP(iterP - 1).nDCI Then
928:                    arraycruz(iterador).nDCI = True
929:                Else
930:                    arraycruz(iterador).nDCI = False
931:                End If

932:                If arrayA(iterA - 1).pvp = arrayP(iterP - 1).pvp Then
933:                    arraycruz(iterador).pvp = True
934:                Else
935:                    arraycruz(iterador).pvp = False
936:                End If

84:                 If arraycruz(iterador).pvp = False Then 'se o pvp não é igual vai ver se é menor, maior    [sem ser unitário]
85:                     If arrayP(iterP - 1).pvp > arrayA(iterA - 1).pvp Then 'se pvp aviado < prescrito
86:                         arraycruz(iterador).xpvp = 1
87:                     ElseIf arrayP(iterP - 1).pvp > arrayA(iterA - 1).pvp Then 'se pvp aviado > prescrito
88:                         arraycruz(iterador).xpvp = 2
89:                     End If
90:                 Else
91:                     arraycruz(iterador).xpvp = 0 'pvp =   
92:                 End If
93:                 If arraycruz(iterador).cnpem = True Then
94:                     If arrayP(iterP - 1).cnpem = 0 Then
95:
96:                         arraycruz(iterador).xcnpem = 3 'sem cnpem
97:                     Else
98:                         arraycruz(iterador).xcnpem = 4 'cnpem é igual
99:                     End If
100:                Else 'é diferente
101:                    If arrayP(iterP - 1).cnpem <> 0 And arrayA(iterA - 1).cnpem <> 0 Then 'se ambos <> 0
102:                        arraycruz(iterador).xcnpem = 0
103:                    ElseIf arrayP(iterP - 1).cnpem = 0 Then 'se prescrito = 0
104:                        arraycruz(iterador).xcnpem = 1
105:                    Else 'aviado =0
106:                        arraycruz(iterador).xcnpem = 2
107:                    End If
108:                End If
109:                If arraycruz(iterador).GH = True Then
110:                    If arrayP(iterP - 1).GH = 0 Then
111:                        arraycruz(iterador).xGH = 3 'sem GH
112:                    Else
113:                        arraycruz(iterador).xGH = 4 'GH é igual
114:                    End If
115:                Else 'é diferente
116:                    If arrayP(iterP - 1).GH <> 0 And arrayA(iterA - 1).GH <> 0 Then 'se ambos <> 0
117:                        arraycruz(iterador).xGH = 0
118:                    ElseIf arrayP(iterP - 1).GH = 0 Then 'se prescrito = 0
119:                        arraycruz(iterador).xGH = 1
120:                    Else 'aviado =0
121:                        arraycruz(iterador).xGH = 2
122:                    End If
123:                End If
124:
125:                If arraycruz(iterador).nDCI = True Then  'é true se ambos iguais
126:                    If arrayP(iterP - 1).nDCI = False Then 'se pelo menos um é falso então é falso
127:                        arraycruz(iterador).xnDCI = False
128:                    Else 'se ambos são verdadeiros é verdadeiro
129:                        arraycruz(iterador).xnDCI = True
130:                    End If
131:                End If
132:
133:                If arrayP(iterP - 1).code > 50000000 Then 'prescrito por CNPEM
134:                    Select Case arraycruz(iterador).xcnpem
                            Case Is = 0 'mesmo cnpem
136:                            arraycruz(iterador).porDCImesmoCNPEM = True
137:                            arraycruz(iterador).porDCIdifCNPEM = False
138:                        Case Is = 1 'prescrito sem cnpem
139:                            arraycruz(iterador).porDCImesmoCNPEM = False
140:                            arraycruz(iterador).porDCIdifCNPEM = True
141:                        Case Is = 2 'aviado sem cnpem
142:                            arraycruz(iterador).porDCImesmoCNPEM = False
143:                            arraycruz(iterador).porDCIdifCNPEM = True
144:                        Case Is = 3 'ambos sem cnpem
145:                            arraycruz(iterador).porDCImesmoCNPEM = True
146:                            arraycruz(iterador).porDCIdifCNPEM = False
147:                        Case Is = 4 'cnpem diferente
148:                            arraycruz(iterador).porDCImesmoCNPEM = False
149:                            arraycruz(iterador).porDCIdifCNPEM = True
150:                    End Select
                        arraycruz(iterador).excepa = 5
                        arraycruz(iterador).excepb = 5
                        arraycruz(iterador).excepc = 5
                        arraycruz(iterador).porDCImesmoCNPEM = 5
151:                Else 'prescrito por marca ou lab (pode ser genérico)
152:                    Select Case arraycruz(iterador).xcnpem
                            Case Is = 0 'mesmo cnpem
154:                            arraycruz(iterador).porMARCAmesmoCNPEM = True
155:                            arraycruz(iterador).porMARCAdifCNPEM = False
156:                        Case Is = 1 'prescrito sem cnpem
157:                            arraycruz(iterador).porMARCAmesmoCNPEM = False
158:                            arraycruz(iterador).porMARCAdifCNPEM = True
159:                        Case Is = 2 'aviado sem cnpem
160:                            arraycruz(iterador).porMARCAmesmoCNPEM = False
161:                            arraycruz(iterador).porMARCAdifCNPEM = True
162:                        Case Is = 3 'ambos sem cnpem
163:                            arraycruz(iterador).porMARCAmesmoCNPEM = True
164:                            arraycruz(iterador).porMARCAdifCNPEM = False
165:                        Case Is = 4 'cnpem diferente
166:                            arraycruz(iterador).porMARCAmesmoCNPEM = False
167:                            arraycruz(iterador).porMARCAdifCNPEM = True
168:                    End Select
                        If arrayP(iterP - 1).excepa = True Then 'se prescrito tem dci dos excep a
171:                        If arraycruz(iterador).code = False Then 'se prescrito e aviado têm codigo diferente
172:                            arraycruz(iterador).excepa = 0
173:                        Else
174:                            arraycruz(iterador).excepa = 1
175:                        End If
176:                    Else 'se o dci não é dos excep a
177:                        arraycruz(iterador).excepa = 2
178:                    End If
179:                    If arraycruz(iterador).code = False Then 'And arraycruz(iterador).lab = False Then 'se prescrito e aviado tÊm nome, lab e código diferentes
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
                        Else 'acrescentado para permitir acusar top5 quando aviado não tem top5 associado, nomeadamente quando é CNPEM
                            If arrayA(iterA - 1).pvp > arrayA(iterA - 1).top5 Then
                                arraycruz(iterador).top5 = 0
                            ElseIf arrayA(iterA - 1).pvp = arrayA(iterA - 1).top5 Then
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
181:                iterP = iterP + 1
182:            Loop
183:            iterA = iterA + 1
184:        Loop
185:
186:
187:        makeranking2013()
188:
189:
190:        Exit Function
MOSTRARERRO:
            MsgBox("F12" & "L" & Str$(Erl) & "E" & Str$(Err.Number))
            Return RedirectToAction("Index", "Input")

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
1:          Select Case LLforma
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
                Case "adesivo transdérmico"
                    via = 204
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
272:            Case "seringa"
273:
                    via = 21
274:            Case "agulhas"
275:
                    via = 22
276:            Case "lancetas"

277:                via = 23
278:            Case "lancetas para punção capilar"

                    via = 23
279:            Case "lancetas para amostragem de sangue"

                    via = 23
280:            Case "lancetas esterilizadas por irradiação, de uso único"

281:                via = 23
283:            Case "lancetas estéreis"

284:                via = 23
285:            Case "lancetas estéreis para obtenção de uma gota de sangue"

286:                via = 23
287:            Case "lancetas estéreis por radiação gama"

288:                via = 23
289:            Case "lancetas esterilizadas por radiação gama"

290:                via = 23
291:            Case "seringa de uso único, estéril para administração de insulina com agulha ultrafina, 30G, 8mm"

292:                via = 21
293:            Case "seringas de 0,3 ml (8mm) diametro 0,3mm (30G) escala de 30 unidades divididas em 1/2"

294:                via = 21
295:            Case "tiras para determinação de glicémia"

296:                via = 24
297:            Case "tiras para determinação de glicosúria"

298:                via = 25
299:            Case "tiras para determinação de glicosúria e cetonúria"

300:                via = 26
301:            Case "tiras teste de b-cetonemia"

302:                via = 27
303:            Case "agulha de 0,30x8mm (30 Gx5/16)"

304:                via = 22
305:            Case "agulha de 0,33x12mm (29 Gx15/32)"

306:                via = 22
307:            Case "agulha de uso único, estéril, para canetas de administração de insulina"

308:                via = 22
309:
509:            Case Else
510:            'MsgBox("forma farmacêutica desconhecida")
511:        End Select
512:        If via = 0 Then
513:        ' MsgBox("via = 0")
514:        End If
515:        Exit Function
MOSTRARERRO:
            MsgBox("F13" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

            Resume Next
        End Function




        Function detectarmaisdedoisaviados() As Boolean
1:          On Error GoTo MOSTRARERRO
2:          Dim aviamento(3) As Object

            '            identificar os que têm mesmo codigo
            'os que têm os mesmo nome com mesmos outros dados    
            'os que têm mesmo cnpem
            'os que têm cnpem = 0 e mesmo outro

            'para ser considerado igual{
            'não é unitário
            'tem mesmo dci, dose, formavia,
            '}


            'se um igual a dois,
            '    se tres é igual then
            '        é erro
            '            se 4 é igual then
            '            tb tem erro
            '            Else
            '            nada
            '            Else se 4 é igual Then
            '            é erro
            '            Else
            '            nada
            '            Else se dois igual a tres
            '    se quatro é igual then
            '        é erro
            '        Else
            '            nada

3:          aviamento(0) = oa1
4:          aviamento(1) = oa2
5:          aviamento(2) = oa3
6:          aviamento(3) = oa4
7:          repeticao = 1
8:          For aviados As Integer = 1 To A - 1
9:              If aviamento(aviados).code = aviamento(aviados - 1).code Then
91:                 If (IsNumeric(aviamento(aviados).qty) AndAlso aviamento(aviados).qty <> 1) Or (antesa(aviados) > 0 And antesa(aviados) <> 1) Then
10:                     repeticao = repeticao + 1
                    End If
11:                 If repeticao = 3 Then
12:                     If detectarmaisdedoisprescritos() Then Return True
13:                     Exit For
14:                 End If

15:             End If
16:         Next
17:
18:         Exit Function
MOSTRARERRO:
            MsgBox("F14" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

            Resume Next
        End Function

        Function detectarmaisdedoisprescritos() As Boolean
1:          On Error GoTo MOSTRARERRO
2:          Dim prescricao(3) As Object
3:          prescricao(0) = op1
4:          prescricao(1) = op2
5:          prescricao(2) = op3
6:          prescricao(3) = op4
7:          repeticao = 1
8:          For prescritos As Integer = 1 To P
9:              If prescricao(prescritos).code = prescricao(prescritos - 1).code Then
                    If (IsNumeric(prescricao(prescritos).qty) AndAlso prescricao(prescritos).qty <> 1) Or (antesa(prescritos) > 0 And antesa(prescritos) <> 1) Then
10:                     repeticao = repeticao + 1
                    End If
11:                 If repeticao = 3 Then
12:                     Return True
13:                     Exit For
14:                 End If
15:             End If
16:         Next
17:
18:         Exit Function
MOSTRARERRO:
            MsgBox("F15" & "L" & Str$(Erl) & "E" & Str$(Err.Number))

            Resume Next
        End Function







    End Class
End Namespace