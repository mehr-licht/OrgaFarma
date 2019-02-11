Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb





Public Class Form2
    Inherits Form




    Private _passedText As String

    Public Property [PassedText]() As String
        Get
            Return _passedText
        End Get
        Set(ByVal Value As String)
            _passedText = Value
        End Set


    End Property

    'colocar verificação = se forem do mesmo GH não compara tamanhos nem apresentações - seria OK não fosse o laboratório

    Dim unidose As Boolean = False
    Dim tectocomp As Short = 95
    Dim manipcomp As Short = 30
    Dim mudarentidadelimpa As Boolean = True
    Dim haport As Boolean = False
    Dim port1474 As Boolean = False
    Dim farmacia As Integer
    Dim tempcalc As Double
    Dim mostrarsubvia As Boolean = True
    Dim mostrarmarca As Boolean = True
    Dim mostrarcoddif As Boolean = False
    Dim mostrartrocalab As Boolean = True
    Dim mostrarqtyinferior As Boolean = False
    Dim taxaQuant As Double = 1.5
    Dim taxapr As Short = 1
    Dim conjunto As Short
    Dim amarelo As Boolean
    Dim vermelho As Boolean
    Dim verde As Boolean
    Dim varcruza1 As Short
    Dim varcruza2 As Short
    Dim varcruza3 As Short
    Dim varcruza4 As Short
    Dim varcruza5 As Short
    Dim varcruza6 As Short
    Dim varcruzp1 As Short
    Dim varcruzp2 As Short
    Dim varcruzp3 As Short
    Dim varcruzp4 As Short
    Dim varcruzp5 As Short
    Dim varcruzp6 As Short
    Dim varlabelcruz1 As String
    Dim varlabelcruz2 As String
    Dim varlabelcruz3 As String
    Dim varlabelcruz4 As String
    Dim varlabelcruz5 As String
    Dim varlabelcruz6 As String
    Dim portcomp As Double
    Dim portcomp01_1 As Double
    Dim portcomp01_2 As Double
    Dim portcomp01_3 As Double
    Dim portcomp01_4 As Double
    Dim portcomp5 As Double
    Dim portcomp6 As Double
    Dim intermedio As Double
    Dim pvp1v As String
    Dim pvp2v As String
    Dim pvp3v As String
    Dim pvp4v As String
    Dim pvp5v As String
    Dim pvp6v As String
    Dim pvp1val As Double
    Dim pvp2val As Double
    Dim pvp3val As Double
    Dim pvp4val As Double
    Dim pvp5val As Double
    Dim pvp6val As Double
    Dim comp01_1v As String
    Dim comp01_2v As String
    Dim comp01_3v As String
    Dim comp01_4v As String
    Dim comp5v As String
    Dim comp6v As String
    Dim comp01_1val As Double
    Dim comp01_2val As Double
    Dim comp01_3val As Double
    Dim comp01_4val As Double
    Dim comp5val As Double
    Dim comp6val As Double
    Dim pr As Double
    Dim pr1 As Double
    Dim pr2 As Double
    Dim pr3 As Double
    Dim pr4 As Double
    Dim pr5 As Double
    Dim pr6 As Double
    Dim organismo As String
    Dim gen As Boolean
    Dim comp As Double
    Dim portimedio As String
    Dim p5row As basededadosDataSet.infarmedRow
    Dim p6row As basededadosDataSet.infarmedRow
    Dim a5row As basededadosDataSet.infarmedRow
    Dim a6row As basededadosDataSet.infarmedRow


    Dim mostrado9 As Boolean
    Dim portec4a As String
    Dim portec4b As String
    Dim genec4 As String
    Dim mostranome As String
    Dim mostradci As String
    Dim mostraforma As String
    Dim mostradoseqty As String
    Dim mostracomp As String
    Dim mostracompgen As String
    Dim mostraports As String
    Dim codigorow As basededadosDataSet.infarmedRow
    Dim codigoarray As New ArrayList
    Dim codigo4 As New meds
    Dim novado As Boolean

    Dim form2 As New bd()

    Dim p1, p2, p3, p4, a1, a2, a3, a4 As Integer

    'trabalhar a data da comparticipação - no exemplo 40023 bate certo com 29/07/2009 - é só substituir por aXarray(5)
    'Dim descomp As Date = DateAdd(DateInterval.Day, 40023, #12/30/1899#)
    'em vez de abrir a msgbox comparar logo com data actual e se fora de data dar como não comparticiapado (¿desde...?)
    Dim descomp01_1mostrado As Boolean
    Dim descomp01_2mostrado As Boolean
    Dim descomp01_3mostrado As Boolean
    Dim descomp01_4mostrado As Boolean

    Dim Prescrito1 As New meds
    Dim Prescrito2 As New meds
    Dim Prescrito3 As New meds
    Dim Prescrito4 As New meds
    Dim Prescrito5 As New meds
    Dim Prescrito6 As New meds

    Dim Aviado1 As New meds
    Dim Aviado2 As New meds
    Dim Aviado3 As New meds
    Dim Aviado4 As New meds
    Dim Aviado5 As New meds
    Dim Aviado6 As New meds
    Dim vazio As New meds

    'usado na avaliação keypress para saber em que controlo está o foco
    Dim foco As String

    'as variaveis A e P contam quantas embalagens foram (A)viadas e quantas foram (P)rescritas
    Dim A As Short
    Dim P As Short

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
    Dim a5array As New ArrayList
    Dim a6array As New ArrayList
    Dim p1array As New ArrayList
    Dim p2array As New ArrayList
    Dim p3array As New ArrayList
    Dim p4array As New ArrayList
    Dim p5array As New ArrayList
    Dim p6array As New ArrayList

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
    Dim av5 As New avaliacao
    Dim av6 As New avaliacao

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


    'Prepara para ler nova receita - limpa tudo (valores a zero, labels em branco e sem fundo) e foco na primeira caixa
    Sub inicializar()
        On Error GoTo MOSTRARERRO
        limpo = True

1:      av1.nivel = 99
2:      av2.nivel = 99
3:      av3.nivel = 99
4:      av4.nivel = 99
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
37:     descomp01_1mostrado = False
38:     descomp01_2mostrado = False
39:     descomp01_3mostrado = False
40:     descomp01_4mostrado = False
41:     A = 0
42:     P = 0
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
71:     Me.presc1.Focus()
72:     Me.Prescrito1.codigo = Nothing
73:     Me.Prescrito2.codigo = Nothing
74:     Me.Prescrito3.codigo = Nothing
75:     Me.Prescrito4.codigo = Nothing
76:     Me.Aviado1.codigo = Nothing
77:     Me.Aviado2.codigo = Nothing
78:     Me.Aviado3.codigo = Nothing
79:     Me.Aviado4.codigo = Nothing
80:     Me.a1p1.nivel = 99
81:     Me.a2p1.nivel = 99
82:     Me.a3p1.nivel = 99
83:     Me.a4p1.nivel = 99
84:     Me.a1p2.nivel = 99
85:     Me.a2p2.nivel = 99
86:     Me.a3p2.nivel = 99
87:     Me.a4p2.nivel = 99
88:     Me.a1p3.nivel = 99
89:     Me.a2p3.nivel = 99
90:     Me.a3p3.nivel = 99
91:     Me.a4p3.nivel = 99
92:     Me.a1p4.nivel = 99
93:     Me.a2p4.nivel = 99
94:     Me.a3p4.nivel = 99
95:     Me.a4p4.nivel = 99
96:     Me.a1p1.resultado = 99
97:     Me.a2p1.resultado = 99
98:     Me.a3p1.resultado = 99
99:     Me.a4p1.resultado = 99
100:    Me.a1p2.resultado = 99
101:    Me.a2p2.resultado = 99
102:    Me.a3p2.resultado = 99
103:    Me.a4p2.resultado = 99
104:    Me.a1p3.resultado = 99
105:    Me.a2p3.resultado = 99
106:    Me.a3p3.resultado = 99
107:    Me.a4p3.resultado = 99
108:    Me.a1p4.resultado = 99
109:    Me.a2p4.resultado = 99
110:    Me.a3p4.resultado = 99
111:    Me.a4p4.resultado = 99
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
156:    varcruza5 = 0
157:    varcruza6 = 0
158:    varcruzp1 = 0
159:    varcruzp2 = 0
160:    varcruzp3 = 0
161:    varcruzp4 = 0
162:    varcruzp5 = 0
163:    varcruzp6 = 0
180:    Me.pvp1.Text = ""
181:    Me.pvp2.Text = ""
182:    Me.pvp3.Text = ""
183:    Me.pvp4.Text = ""
186:    Me.comp01_1.Text = ""
187:    Me.comp01_2.Text = ""
188:    Me.comp01_3.Text = ""
189:    Me.comp01_4.Text = ""
192:    Me.pvp_tot.Text = ""
193:    Me.comp01_tot.Text = ""
195:    verde = True
196:    amarelo = False
197:    vermelho = False
        conjunto = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB INICIALIZAR: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Sub LimparRes()
        On Error GoTo MOSTRARERRO
        limpo = True
1:      av1.nivel = 99
2:      av2.nivel = 99
3:      av3.nivel = 99
4:      av4.nivel = 99
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
37:     descomp01_1mostrado = False
38:     descomp01_2mostrado = False
39:     descomp01_3mostrado = False
40:     descomp01_4mostrado = False
41:     Me.aviam1.Text = ""
42:     Me.aviam2.Text = ""
43:     Me.aviam3.Text = ""
44:     Me.aviam4.Text = ""
49:     Me.aviam1.BackColor = Color.White
50:     Me.aviam2.BackColor = Color.White
51:     Me.aviam3.BackColor = Color.White
52:     Me.aviam4.BackColor = Color.White
57:     Me.a1p1.nivel = 99
58:     Me.a2p1.nivel = 99
59:     Me.a3p1.nivel = 99
60:     Me.a4p1.nivel = 99
61:     Me.a1p2.nivel = 99
62:     Me.a2p2.nivel = 99
63:     Me.a3p2.nivel = 99
64:     Me.a4p2.nivel = 99
65:     Me.a1p3.nivel = 99
66:     Me.a2p3.nivel = 99
67:     Me.a3p3.nivel = 99
68:     Me.a4p3.nivel = 99
69:     Me.a1p4.nivel = 99
70:     Me.a2p4.nivel = 99
71:     Me.a3p4.nivel = 99
72:     Me.a4p4.nivel = 99
73:     Me.a1p1.resultado = 99
74:     Me.a2p1.resultado = 99
75:     Me.a3p1.resultado = 99
76:     Me.a4p1.resultado = 99
77:     Me.a1p2.resultado = 99
78:     Me.a2p2.resultado = 99
79:     Me.a3p2.resultado = 99
80:     Me.a4p2.resultado = 99
81:     Me.a1p3.resultado = 99
82:     Me.a2p3.resultado = 99
83:     Me.a3p3.resultado = 99
84:     Me.a4p3.resultado = 99
85:     Me.a1p4.resultado = 99
86:     Me.a2p4.resultado = 99
87:     Me.a3p4.resultado = 99
88:     Me.a4p4.resultado = 99
89:     Me.aviam1.Focus()
90:     Me.a1p1.mostrado = False
91:     Me.a2p1.mostrado = False
92:     Me.a3p1.mostrado = False
93:     Me.a4p1.mostrado = False
94:     Me.a1p2.mostrado = False
95:     Me.a2p2.mostrado = False
96:     Me.a3p2.mostrado = False
97:     Me.a4p2.mostrado = False
98:     Me.a1p3.mostrado = False
99:     Me.a2p3.mostrado = False
100:    Me.a3p3.mostrado = False
101:    Me.a4p3.mostrado = False
102:    Me.a1p4.mostrado = False
103:    Me.a2p4.mostrado = False
104:    Me.a3p4.mostrado = False
105:    Me.a4p4.mostrado = False
106:    grupoP1 = 0
107:    grupoP2 = 0
108:    grupoP3 = 0
109:    grupoP4 = 0
110:    grupoA1 = 0
111:    grupoA2 = 0
112:    grupoA3 = 0
113:    grupoA4 = 0
114:    grupoP1dci = ""
115:    grupoP2dci = ""
116:    grupoP3dci = ""
117:    grupoP4dci = ""
118:    grupoA1dci = ""
119:    grupoA2dci = ""
120:    grupoA3dci = ""
121:    grupoA4dci = ""
152:    varcruza1 = 0
153:    varcruza2 = 0
154:    varcruza3 = 0
155:    varcruza4 = 0
156:    varcruza5 = 0
157:    varcruza6 = 0
158:    varcruzp1 = 0
159:    varcruzp2 = 0
160:    varcruzp3 = 0
161:    varcruzp4 = 0
162:    varcruzp5 = 0
163:    varcruzp6 = 0

180:    Me.pvp1.Text = ""
181:    Me.pvp2.Text = ""
182:    Me.pvp3.Text = ""
183:    Me.pvp4.Text = ""

186:    Me.comp01_1.Text = ""
187:    Me.comp01_2.Text = ""
188:    Me.comp01_3.Text = ""
189:    Me.comp01_4.Text = ""
192:    Me.pvp_tot.Text = ""
193:    Me.comp01_tot.Text = ""

195:    verde = True
196:    amarelo = False
197:    vermelho = False

        conjunto = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB LIMPARRES: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub LimparPrescritos()
        On Error GoTo MOSTRARERRO
        limpo = True
1:      av1.nivel = 99
2:      av2.nivel = 99
3:      av3.nivel = 99
4:      av4.nivel = 99
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
37:     Me.presc1.Text = ""
38:     Me.presc2.Text = ""
39:     Me.presc3.Text = ""
40:     Me.presc4.Text = ""
45:     Me.a1p1.mostrado = False
46:     Me.a2p1.mostrado = False
47:     Me.a3p1.mostrado = False
48:     Me.a4p1.mostrado = False
49:     Me.a1p2.mostrado = False
50:     Me.a2p2.mostrado = False
51:     Me.a3p2.mostrado = False
52:     Me.a4p2.mostrado = False
53:     Me.a1p3.mostrado = False
54:     Me.a2p3.mostrado = False
55:     Me.a3p3.mostrado = False
56:     Me.a4p3.mostrado = False
57:     Me.a1p4.mostrado = False
58:     Me.a2p4.mostrado = False
59:     Me.a3p4.mostrado = False
60:     Me.a4p4.mostrado = False
61:     Me.presc1.BackColor = Color.White
62:     Me.presc2.BackColor = Color.White
63:     Me.presc3.BackColor = Color.White
64:     Me.presc4.BackColor = Color.White
69:     Me.a1p1.nivel = 99
70:     Me.a2p1.nivel = 99
71:     Me.a3p1.nivel = 99
72:     Me.a4p1.nivel = 99
73:     Me.a1p2.nivel = 99
74:     Me.a2p2.nivel = 99
75:     Me.a3p2.nivel = 99
76:     Me.a4p2.nivel = 99
77:     Me.a1p3.nivel = 99
78:     Me.a2p3.nivel = 99
79:     Me.a3p3.nivel = 99
80:     Me.a4p3.nivel = 99
81:     Me.a1p4.nivel = 99
82:     Me.a2p4.nivel = 99
83:     Me.a3p4.nivel = 99
84:     Me.a4p4.nivel = 99
85:     Me.presc1.Focus()
94:     grupoP1 = 0
95:     grupoP2 = 0
96:     grupoP3 = 0
97:     grupoP4 = 0
98:     grupoA1 = 0
99:     grupoA2 = 0
100:    grupoA3 = 0
101:    grupoA4 = 0
102:    grupoP1dci = ""
103:    grupoP2dci = ""
104:    grupoP3dci = ""
105:    grupoP4dci = ""
106:    grupoA1dci = ""
107:    grupoA2dci = ""
108:    grupoA3dci = ""
109:    grupoA4dci = ""
152:    varcruza1 = 0
153:    varcruza2 = 0
154:    varcruza3 = 0
155:    varcruza4 = 0
156:    varcruza5 = 0
157:    varcruza6 = 0
158:    varcruzp1 = 0
159:    varcruzp2 = 0
160:    varcruzp3 = 0
161:    varcruzp4 = 0
162:    varcruzp5 = 0
163:    varcruzp6 = 0
        conjunto = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB LIMPARPRESCRITOS: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao clickar no botão limpar prescritos
    Private Sub limpresc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      LimparPrescritos()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limprec_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'o que acontece ao fazer enter no botão limpar prescritos
    Private Sub limpresc_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      LimparPrescritos()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limpresc_Enter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao fazer enter no botão limpar aviados (e resultados)
    'Private Sub limpar_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles limpar.Enter
    '   On Error GoTo MOSTRARERRO
    '  LimparRes()
    ' Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub limpar_Enter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub


    'o que acontece ao clickar no botão limpar aviados (e resultados)
    Private Sub limpar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      LimparRes()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limpar_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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



    'o que acontece ao abrir/inicializar o form principal - inicia timer, limpa tudo, poe tudo a zero e foca na 1ª textbox
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'TODO: This line of code loads data into the 'BasededadosDataSet1.infarmed' table. You can move, or remove it, as needed.

        farmacia = Convert.ToInt32(_passedText)

        'Me.InfarmedTableAdapter1.Fill(Me.BasededadosDataSet1.infarmed)
        On Error GoTo MOSTRARERRO
1:      'TODO: This line of code loads data into the 'EscolhaDataSet22.lab' table. You can move, or remove it, as needed.
2:      'Me.LabTableAdapter3.Fill(Me.EscolhaDataSet22.lab)
3:      'TODO: This line of code loads data into the 'EscolhaDataSet21.lab' table. You can move, or remove it, as needed.
4:      'Me.LabTableAdapter2.Fill(Me.EscolhaDataSet21.lab)
5:      'TODO: This line of code loads data into the 'EscolhaDataSet20.lab' table. You can move, or remove it, as needed.
6:      'Me.LabTableAdapter.Fill(Me.EscolhaDataSet20.lab)
7:      'TODO: This line of code loads data into the 'EscolhaDataSet19.qty' table. You can move, or remove it, as needed.
8:      'Me.QtyTableAdapter3.Fill(Me.EscolhaDataSet19.qty)
9:      'TODO: This line of code loads data into the 'EscolhaDataSet18.qty' table. You can move, or remove it, as needed.
10:     'Me.QtyTableAdapter2.Fill(Me.EscolhaDataSet18.qty)
11:     'TODO: This line of code loads data into the 'EscolhaDataSet17.qty' table. You can move, or remove it, as needed.
12:     'Me.QtyTableAdapter.Fill(Me.EscolhaDataSet17.qty)
13:     'TODO: This line of code loads data into the 'EscolhaDataSet16.forma' table. You can move, or remove it, as needed.
14:     'Me.FormaTableAdapter3.Fill(Me.EscolhaDataSet16.forma)
15:     'TODO: This line of code loads data into the 'EscolhaDataSet15.forma' table. You can move, or remove it, as needed.
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
26:     'Me.Dc1TableAdapter1.Fill(Me.EscolhaDataSet10.dc1)
27:     'TODO: This line of code loads data into the 'EscolhaDataSet9.lab' table. You can move, or remove it, as needed.
28:     'Me.LabTableAdapter1.Fill(Me.EscolhaDataSet9.lab)
29:     'TODO: This line of code loads data into the 'EscolhaDataSet8.qty' table. You can move, or remove it, as needed.
30:     'Me.QtyTableAdapter1.Fill(Me.EscolhaDataSet8.qty)
31:     'TODO: This line of code loads data into the 'EscolhaDataSet7.dose' table. You can move, or remove it, as needed.
32:     'Me.DoseTableAdapter1.Fill(Me.EscolhaDataSet7.dose)
33:     'TODO: This line of code loads data into the 'EscolhaDataSet6.forma' table. You can move, or remove it, as needed.
34:     'Me.FormaTableAdapter1.Fill(Me.EscolhaDataSet6.forma)
35:     'TODO: This line of code loads data into the 'EscolhaDataSet5.dc1' table. You can move, or remove it, as needed.
36:     'Me.Dc1TableAdapter.Fill(Me.EscolhaDataSet5.dc1)
37:     'Me.InfarmedTableAdapter.Fill(Me.BasededadosDataSet.infarmed)
38:
39:     inicializar()
        Me.WindowState = FormWindowState.Maximized
40:     Me.KeyPreview = True

        organismo = "01"

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Form1_Load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
    '2:      codEC.Text = ""
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
8:
9:      Select Case soalgarismos
            Case False
11:             Beep()
12:
        End Select
13:     Select Case foco
            Case "presc1"
15:             Beep()
17:             presc1.Text = "0"
18:
19:         Case "presc2"
20:
21:             Beep()
22:             presc2.Text = "0"
23:
24:         Case "presc3"
25:
26:             Beep()
27:             presc3.Text = "0"

29:         Case "presc4"

31:             Beep()
32:             presc4.Text = "0"

a33:
34:             End Select
35:
36:
37:
38:     If Asc(e.KeyChar) = Keys.Enter Then
39:         Select Case foco
                Case "presc1"
55:                 If presc1.Text = "" Then
56:                     Beep()
57:                 Else
58:                     Me.presc2.Focus()
59:                 End If
60:
61:             Case "aviam1"
62:                 If aviam1.Text = "" Then
63:                     Beep()
64:                 Else
65:                     Me.aviam2.Focus()
66:                 End If
67:             Case "presc2"
68:                 If presc2.Text = "" Then
69:                     Me.aviam1.Focus()
70:                     Me.presc2.Text = "0"
71:                     Me.presc3.Text = "0"
72:                     Me.presc4.Text = "0"
a72:
73:                 Else
74:                     Me.presc3.Focus()
75:                 End If
76:             Case "presc3"
77:                 If presc3.Text = "" Then
78:                     Me.aviam1.Focus()
79:                     Me.presc3.Text = "0"
80:                     Me.presc4.Text = "0"
a80:
81:                 Else
82:                     Me.presc4.Focus()
83:                 End If
84:             Case "presc4"
85:                 If presc4.Text = "" Then
86:                     Me.aviam1.Focus()
87:                     Me.presc4.Text = "0"
a87:
88:                 Else
89:                     Me.aviam1.Focus()

                    End If

91:             Case "aviam2"
92:                 If aviam2.Text = "" Then
93:                     Me.aviam2.Text = "0"
94:                     Me.aviam3.Text = "0"
95:                     Me.aviam4.Text = "0"
a95:
96:                     Comparar()
a96:                    somas()
97:                 Else
98:                     Me.aviam3.Focus()
99:                 End If
100:            Case "aviam3"
101:                If aviam3.Text = "" Then
102:                    Me.aviam3.Text = "0"
103:                    Me.aviam4.Text = "0"
a103:
104:                    Comparar()
a104:                   somas()
105:                Else
106:                    Me.aviam4.Focus()
107:                End If
a107:           Case "aviam4"
z92:                If aviam4.Text = "" Then
z93:                    Me.aviam4.Text = "0"
z94:
                        Comparar()
zzz96:                  somas()
                        Me.presc1.Focus()
z96:                Else

b110:                   If aviam4.Text >= 1111111 And aviam4.Text <= 9999999 Then
c110:                       Comparar()
d110:                       somas()
e110:                       Me.presc1.Focus()
f110:                   Else
g110:                       Beep()
h110:                       Me.aviam4.Text = ""
i110:                       Me.aviam4.Focus()
z99:                    End If
                    End If



z100:
112:                Comparar()
                    somas()
113:                End Select
114:    End If
115:
116:
117:    If Asc(e.KeyChar) = Keys.C Then
118:        My.Computer.Keyboard.SendKeys("{bs}")
119:        My.Computer.Keyboard.SendKeys("{ENTER}")
120:    End If
121:
122:    If Asc(e.KeyChar) = Keys.F1 Then
123:        inicializar()
124:    End If
125:
126:    If Asc(e.KeyChar) = Keys.F12 Then
127:        Comparar()
128:    End If
129:
130:    If Asc(e.KeyChar) = Keys.F2 Then
131:        form2.Show()
132:    End If
133:
134:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.I Then
135:        inicializar()
136:    End If
137:
138:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.C Then
139:        Comparar()
140:    End If
141:
142:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.B Then
143:        form2.Show()
144:    End If
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
176:    If Asc(e.KeyChar) = Keys.Space Then
177:        inicializar()
            limpar4()

178:        Me.Focus()
179:        inicializar()
180:    End If
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
4:      If Len(presc1.Text) = 7 Then
5:          If Caractp1 Like "#######" Then
6:              Prescrito1.codigo = presc1.Text
7:              'Me.presc2.Focus()
8:          End If
9:
10:     Else
11:         presc1.Text = "0"
12:     End If
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
4:      If Len(presc2.Text) = 7 Then
5:          If Caractp2 Like "#######" Then
6:              Prescrito2.codigo = presc2.Text
7:              'Me.presc3.Focus()
8:          End If
9:
10:     Else
11:         presc2.Text = "0"
12:     End If
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
4:      If Len(presc3.Text) = 7 Then
5:          If Caractp3 Like "#######" Then
6:              Prescrito3.codigo = presc3.Text
7:              'Me.presc4.Focus()
8:          End If
9:
10:     Else
11:         presc3.Text = "0"
12:     End If
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
4:      If Len(presc4.Text) = 7 Then
5:          If Caractp4 Like "#######" Then
6:              Prescrito4.codigo = presc4.Text
7:              'Me.aviam1.Focus()
8:          End If
9:
10:     Else
11:         presc4.Text = "0"
12:     End If
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
                av1.mostrado = "True"
                A = 1
                a1row = DS.infarmed.FindBycode(Aviado1.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado1.codigo)
                a1array.Add(a1row)
                indicar(1)
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
                av2.mostrado = "True"
                A = 2
                a2row = DS.infarmed.FindBycode(Aviado2.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado2.codigo)
                a2array.Add(a2row)
                indicar(2)
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
                av3.mostrado = "True"
                A = 3
                a3row = DS.infarmed.FindBycode(Aviado3.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado3.codigo)
                a3array.Add(a3row)
                indicar(3)
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
                av4.mostrado = "True"
                A = 4
                a4row = DS.infarmed.FindBycode(Aviado4.codigo)
                codigorow = DS.infarmed.FindBycode(Aviado4.codigo)
                a4array.Add(a4row)
                indicar(4)
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





    'chama o irbuscar() e o row2array para produzir os resultados
    Private Sub Comparar()
        On Error GoTo MOSTRARERRO
1:
9:      irbuscar()
10:     'row2array()
11:     TresPresc()
16:     limparzeros()
18:
        'saltar wrapitup para funcionar como v3.0
        'wrapitup()
        limpo = False

20:     'MsgBox("a3p2.nivel = " & a3p2.nivel & vbCr & "a4p2.nivel = " & a4p2.nivel)
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub Comparar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
19:

        Exit Sub
MOSTRARERRO:
        MsgBox("Sub limparzeros: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'vai buscar à baseDeDados as linhas correspondentes aos códigos introduzidos nas 8 caixas
    Sub irbuscar()
1:      On Error GoTo MOSTRARERRO
2:      'AviadosTable = Nothing
3:      'PrescritosTable = Nothing
4:
5:      'era assim antes de eu retirar a verificação se estava escolhida a unidose
        'If aviam4.Text <> "0" Or semcod4.Checked = True And aviam4.Text <> "" Then
6:      If aviam4.Text <> "0" And aviam4.Text <> "" Then
7:          A = 4
            If aviam4.Text = " " Then
                A = 3
            End If
8:      ElseIf aviam3.Text <> "0" And aviam3.Text <> "" Then
9:          A = 3
            If aviam3.Text = " " Then
                A = 2
            End If
10:     ElseIf aviam2.Text <> "0" And aviam2.Text <> "" Then
11:         A = 2
            If aviam2.Text = " " Then
                A = 1
            End If

12:     Else : A = 1
13:     End If
14:
15:     If presc4.Text <> "0" And presc4.Text <> "" Then
16:         P = 4
            If presc4.Text = " " Then
                P = 3
            End If
17:     ElseIf presc3.Text <> "0" And presc3.Text <> "" Then
18:         P = 3
            If presc3.Text = " " Then
                P = 2
            End If
19:     ElseIf presc2.Text <> "0" And presc2.Text <> "" Then
20:         P = 2
            If presc2.Text = " " Then
                P = 1
            End If
21:     Else : P = 1
22:     End If
23:
24:
25:
26:
27:     Select Case A
            Case Is = 1
29:
30:             a1row = DS.infarmed.FindBycode(Aviado1.codigo)
31:             a1array.Add(a1row)
32:             ' Else
33:             '    If presc1dci Then
34:             'a1array(1) = presc1dci
35:             'A1array(2) = presc1forma
36:             'a1array(3) = presc1dose
37:             'a1array(4) = presc1qty
38:             'a1array(8) = presc1lab
40:             'a1array(7) = presc1gen
41:             'End If
42:
70:
71:             On Error Resume Next
72:             Dim conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\ficheiro1.mdb")
73:             Dim cmd As New OleDbCommand("SELECT * INTO [um] FROM [Text;Database=c:\;Hdr=No].[ficheiro1.txt]", conn)
74:             conn.Open()
75:             cmd.ExecuteNonQuery()
76:             conn.Close()
77:             On Error GoTo MOSTRARERRO
78:             'Me.umTableAdapter.Fill(Me.ficheiro1DataSet.um)
79:             'a1row = FS.um.FindBycode("ï»¿9999")
80:             a1array.Add(a1row)
81:
82:
83:             '    If presc1dci Then
84:             'a1array(1) = presc1dci
85:             'A1array(2) = presc1forma
86:             'a1array(3) = presc1dose
87:             'a1array(4) = presc1qty
88:             'a1array(8) = presc1lab
89:             'a1array(7) = presc1gen
90:             'End If
91:
92:
93:             'a1array.Add(linha1)
94:
95:         Case Is = 2
96:
97:
98:             a1row = DS.infarmed.FindBycode(Aviado1.codigo)
99:             a1array.Add(a1row)
100:
101:
102:            a1array.Add(linha1)
103:
104:
105:            a2row = DS.infarmed.FindBycode(Aviado2.codigo)
106:            a2array.Add(a2row)
107:
110:        Case Is = 3
111:
112:
113:            a1row = DS.infarmed.FindBycode(Aviado1.codigo)
114:            a1array.Add(a1row)
115:
119:
120:            a2row = DS.infarmed.FindBycode(Aviado2.codigo)
121:            a2array.Add(a2row)
122:
126:            a3row = DS.infarmed.FindBycode(Aviado3.codigo)
127:            a3array.Add(a3row)
128:
131:        Case Is = 4
132:
133:
134:            a1row = DS.infarmed.FindBycode(Aviado1.codigo)
135:            a1array.Add(a1row)
136:
140:            a2row = DS.infarmed.FindBycode(Aviado2.codigo)
141:            a2array.Add(a2row)
142:
146:            a3row = DS.infarmed.FindBycode(Aviado3.codigo)
147:            a3array.Add(a3row)
148:
152:            a4row = DS.infarmed.FindBycode(Aviado4.codigo)
153:            a4array.Add(a4row)
154:
157:            End Select
158:
159:
160:    Select Case P
            Case Is = 1
162:            p1row = DS.infarmed.FindBycode(Prescrito1.codigo)
163:            p1array.Add(p1row)
164:        Case Is = 2
165:            p1row = DS.infarmed.FindBycode(Prescrito1.codigo)
166:            p2row = DS.infarmed.FindBycode(Prescrito2.codigo)
167:            p1array.Add(p1row)
168:            p2array.Add(p2row)
169:        Case Is = 3
170:            p1row = DS.infarmed.FindBycode(Prescrito1.codigo)
171:            p2row = DS.infarmed.FindBycode(Prescrito2.codigo)
172:            p3row = DS.infarmed.FindBycode(Prescrito3.codigo)
173:            p1array.Add(p1row)
174:            p2array.Add(p2row)
175:            p3array.Add(p3row)
176:        Case Is = 4
177:            p1row = DS.infarmed.FindBycode(Prescrito1.codigo)
178:            p2row = DS.infarmed.FindBycode(Prescrito2.codigo)
179:            p3row = DS.infarmed.FindBycode(Prescrito3.codigo)
180:            p4row = DS.infarmed.FindBycode(Prescrito4.codigo)
181:            p1array.Add(p1row)
182:            p2array.Add(p2row)
183:            p3array.Add(p3row)
184:            p4array.Add(p4row)
185:            End Select
186:    Exit Sub
MOSTRARERRO:
        MsgBox("Sub irbuscar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Public Sub atribuirp()
        On Error GoTo MOSTRARERRO
1:      If P >= 4 Then
2:          Prescrito4.principio = p4row(1)
3:          Prescrito4.apresentacao = p4row(2)
4:          Prescrito4.dosagem = p4row(3)
5:          Prescrito4.quantidade = p4row(4)
6:          Prescrito4.comparticipacao = p4row(5)
7:          Prescrito4.grupo = p4row(6)
8:          Prescrito4.generico = p4row(7)
9:          Prescrito4.laboratorio = p4row(8)
10:     ElseIf P >= 3 Then
11:         Prescrito4 = vazio
12:         Prescrito3.principio = p3row(1)
13:         Prescrito3.apresentacao = p3row(2)
14:         Prescrito3.dosagem = p3row(3)
15:         Prescrito3.quantidade = p3row(4)
16:         Prescrito3.comparticipacao = p3row(5)
17:         Prescrito3.grupo = p3row(6)
18:         Prescrito3.generico = p3row(7)
19:         Prescrito3.laboratorio = p3row(8)
20:
21:     ElseIf P >= 2 Then
22:         Prescrito4 = vazio
23:         Prescrito3 = vazio
24:         Prescrito2.principio = p2row(1)
25:         Prescrito2.apresentacao = p2row(2)
26:         Prescrito2.dosagem = p2row(3)
27:         Prescrito2.quantidade = p2row(4)
28:         Prescrito2.comparticipacao = p2row(5)
29:         Prescrito2.grupo = p2row(6)
30:         Prescrito2.generico = p2row(7)
31:         Prescrito2.laboratorio = p2row(8)
32:
33:     Else
34:         Prescrito4 = vazio
35:         Prescrito3 = vazio
36:         Prescrito2 = vazio
37:         Prescrito1.principio = p1row(1)
38:         Prescrito1.apresentacao = p1row(2)
39:         Prescrito1.dosagem = p1row(3)
40:         Prescrito1.quantidade = p1row(4)
41:         Prescrito1.comparticipacao = p1row(5)
42:         Prescrito1.grupo = p1row(6)
43:         Prescrito1.generico = p1row(7)
44:         Prescrito1.laboratorio = p1row(8)
45:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub atribuirp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Public Sub atribuira()
        On Error GoTo MOSTRARERRO
1:      If A >= 4 Then
2:          Aviado4.principio = a4row(1)
3:          Aviado4.apresentacao = a4row(2)
4:          Aviado4.dosagem = a4row(3)
5:          Aviado4.quantidade = a4row(4)
6:          Aviado4.comparticipacao = a4row(5)
7:          Aviado4.grupo = a4row(6)
8:          Aviado4.generico = a4row(7)
9:          Aviado4.laboratorio = a4row(8)
10:     ElseIf A >= 3 Then
11:         Aviado4 = vazio
12:         Aviado3.principio = a3row(1)
13:         Aviado3.apresentacao = a3row(2)
14:         Aviado3.dosagem = a3row(3)
15:         Aviado3.quantidade = a3row(4)
16:         Aviado3.comparticipacao = a3row(5)
17:         Aviado3.grupo = a3row(6)
18:         Aviado3.generico = a3row(7)
19:         Aviado3.laboratorio = a3row(8)
20:     ElseIf A >= 2 Then
21:         Aviado4 = vazio
22:         Aviado3 = vazio
23:         Aviado2.principio = a2row(1)
24:         Aviado2.apresentacao = a2row(2)
25:         Aviado2.dosagem = a2row(3)
26:         Aviado2.quantidade = a2row(4)
27:         Aviado2.comparticipacao = a2row(5)
28:         Aviado2.grupo = a2row(6)
29:         Aviado2.generico = a2row(7)
30:         Aviado2.laboratorio = a2row(8)
31:     Else
32:         Aviado4 = vazio
33:         Aviado3 = vazio
34:         Aviado2 = vazio
35:         Aviado1.principio = a1row(1)
36:         Aviado1.apresentacao = a1row(2)
37:         Aviado1.dosagem = a1row(3)
38:         Aviado1.quantidade = a1row(4)
39:         Aviado1.comparticipacao = a1row(5)
40:         Aviado1.grupo = a1row(6)
41:         Aviado1.generico = a1row(7)
42:         Aviado1.laboratorio = a1row(8)
43:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub atribuira: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Dim linha1(0 To 8)
    Dim linha2(0 To 8)
    Dim linha3(0 To 8)
    Dim linha4(0 To 8)

    'a1array.Add(linha1) - colocar isto no if semcod1.checked=true
    'dá erro. corrigir e transpor para os outros aviam's e os outros subaviam's
    'em todas as verificações de prescX.text e de pXrow(0) e pXarray(0) tem de se colocar antes o if semcodX.checked=True



 

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






    Sub TresPresc()
        On Error GoTo MOSTRARERRO
1:      If grupoP1 >= 3 Then
2:          MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
3:      End If
4:      If grupoP2 >= 3 Then
5:          MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
6:      End If
7:      If grupoP3 >= 3 Then
8:          MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
9:      End If
10:     If grupoP4 >= 3 Then
11:         MsgBox("Prescritas mais de três embalagens com o mesmo DCI")
12:     End If
13:     If grupoA1 >= 3 Then
14:         MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
15:     End If
16:     If grupoA2 >= 3 Then
17:         MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
18:     End If
19:     If grupoA3 >= 3 Then
20:         MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
21:     End If
22:     If grupoA4 >= 3 Then
23:         MsgBox("Aviadas mais de três embalagens com o mesmo DCI")
24:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'a troca de apresentação é avaliada em função da de administração (acrescentei depois o subvia que avalia a forma mesmo).
    'cada uma tem muitas formas. aqui se associa a via à forma




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





    Sub limpar4()
        On Error GoTo MOSTRARERRO
1:      mostrado9 = "false"
2:      portec4a = ""
3:      portec4b = ""
4:      genec4 = ""
5:      mostranome = ""
6:      mostradci = ""
7:      mostraforma = ""
8:      mostradoseqty = ""
9:      mostracomp = ""
10:     mostracompgen = ""
11:     mostraports = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub incorporar()
        On Error GoTo MOSTRARERRO
1:      If novado = "false" Then
2:          irbuscar4()
3:      End If
4:
5:      If Not IsNothing(codigorow) Then
6:          If codigorow(7) = True Then
7:              genec4 = "genérico"
8:          Else
9:              genec4 = "de marca"
10:         End If
11:
12:         If codigorow(9) = True Then
13:             portec4a = "despacho 4250/2007"
14:             portec4b = ""
15:         Else
16:             If codigorow(11) = True Then
17:                 portec4a = "despacho 10279/2008"
18:                 portec4b = "despacho 10280/2008"
19:             Else
20:                 If codigorow(15) = True Then
                        If port1474 = True Then
21:                         portec4a = "portaria 1474/2004 (ad)"
22:                         portec4b = ""
                        End If
23:                 Else
24:                     If codigorow(20) = True Then
                            If port1474 = True Then
25:                             portec4a = "portaria 1474/2004 (nl)"
26:                             portec4b = ""
                            End If
27:                     Else
28:                         If codigorow(19) = True Then
29:                             portec4a = "despacho 21094/1999"
30:                             portec4b = ""
31:                         Else
32:                             If codigorow(13) = True Then
33:                                 portec4a = "despacho 10910/2009"
34:                                 portec4b = ""
35:                             Else
a36:                                If codigorow(21) = True Then
a37:                                    portec4a = "lei 6/2010"
a38:                                    portec4b = ""
a39:                                Else
36:                                     If codigorow(10) = True Then
37:                                         portec4a = "despacho 1234/2007"
38:                                         If codigorow(14) = True Then
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
51:
57:         mostranome = (codigorow(16))
58:         mostradci = (codigorow(1))
59:         mostraforma = (codigorow(2))
60:         mostradoseqty = "(" & (codigorow(3)) & ") (" & (codigorow(4)) & ")"
61:         mostracompgen = (codigorow(5)) & "% (" & (genec4) & ") "
62:         mostraports = "(" & (portec4a) & ") (" & (portec4b) & ")"
63:         If novado = "false" Then
64:         End If
65:     ElseIf mostrado9 = False Then
66:
67:         mostrado9 = "True"
68:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub irbuscar4()
        On Error GoTo MOSTRARERRO
1:      codigorow = DS.infarmed.FindBycode(codigo4.codigo)
2:      If Not IsNothing(codigorow) Then
3:          codigoarray.Add(codigorow)
4:
10:     ElseIf mostrado9 = False Then
11:
12:         mostrado9 = "True"
13:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub mostrar()
        On Error GoTo MOSTRARERRO
1:      If mostrado9 = False Then
2:
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    '    Private Sub limparEC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        On Error GoTo MOSTRARERRO
    '1:      limpar4()
    '2:      codEC.Text = ""
    '        Exit Sub
    'MOSTRARERRO:
    '        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '        Resume Next
    '    End Sub


    Sub indicar(ByVal which As Short)
1:      On Error GoTo MOSTRARERRO
2:      If Not IsNothing(codigorow) Then
3:          Select Case which
                Case 1
4:                  If av1.mostrado = "true" Then
5:                      If a1row(7) = "true" Then
6:
7:                          gen = True
8:                      Else
9:
10:                         gen = False
11:                     End If
12:                     portaria()
13:
14:                     comp = (a1row(5) * 0.01)
15:                     portcomp01_1 = portcomp
16:                     intermedio = Replace(a1row(17), ".", ",")
17:                     pvp1.Text = intermedio
18:                     pr1 = Replace(a1row(18), ".", ",")
19:                     If organismo = 48 Or organismo = 49 Then
20:                         pr1 = taxapr * pr1
21:                     End If
22:                     pr = pr1
23:                     If pr > 0 Then
                            intermedio = pr
24:                         'o de baixo era quando era usado PVP se PVP<PR
                            'intermedio = System.Math.Min(intermedio, pr)
25:                     End If
                        tempcalc = Replace(a1row(17), ".", ",")
26:                     comp01_1.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                        'comp01_1.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
27:                 End If
28:             Case 2
29:                 If av2.mostrado = "true" Then
30:                     If a2row(7) = "true" Then
31:
32:                         gen = True
33:                     Else
34:
35:                         gen = False
36:                     End If
37:                     portaria()
38:
39:                     comp = (a2row(5) * 0.01)
40:                     portcomp01_2 = portcomp
41:                     intermedio = Replace(a2row(17), ".", ",")
42:                     pvp2.Text = intermedio
43:                     pr2 = Replace(a2row(18), ".", ",")
44:                     If organismo = 48 Or organismo = 49 Then
45:                         pr2 = taxapr * pr2
46:                     End If
47:                     pr = pr2
48:                     If pr > 0 Then
                            intermedio = pr
49:                         'o de baixo era quando era usado PVP se PVP<PR
                            'intermedio = System.Math.Min(intermedio, pr)
50:                     End If
                        tempcalc = Replace(a2row(17), ".", ",")
51:                     comp01_2.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                        'comp01_2.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
52:                 End If
53:             Case 3
54:                 If av3.mostrado = "true" Then
55:                     If a3row(7) = "true" Then
56:
57:                         gen = True
58:                     Else
59:
60:                         gen = False
61:                     End If
62:                     portaria()
63:
64:                     comp = (a3row(5) * 0.01)
65:                     portcomp01_3 = portcomp
66:                     intermedio = Replace(a3row(17), ".", ",")
67:                     pvp3.Text = intermedio
68:                     pr3 = Replace(a3row(18), ".", ",")
69:                     If organismo = 48 Or organismo = 49 Then
70:                         pr3 = taxapr * pr3
71:                     End If
72:                     pr = pr3
73:                     If pr > 0 Then
                            intermedio = pr
                            'o de baixo era quando era usado PVP se PVP<PR
74:                         'intermedio = System.Math.Min(intermedio, pr)
75:                     End If
                        tempcalc = Replace(a3row(17), ".", ",")
76:                     comp01_3.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                        'comp01_3.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
77:                 End If
78:             Case 4
79:                 If av4.mostrado = "true" Then
80:                     If a4row(7) = "true" Then
81:
82:                         gen = True
83:                     Else
84:
85:                         gen = False
86:                     End If
87:                     portaria()

89:                     comp = (a4row(5) * 0.01)
90:                     portcomp01_4 = portcomp
91:                     intermedio = Replace(a4row(17), ".", ",")
92:                     pvp4.Text = intermedio
93:                     pr4 = Replace(a4row(18), ".", ",")
94:                     If organismo = 48 Or organismo = 49 Then
95:                         pr4 = taxapr * pr4
96:                     End If
97:                     pr = pr4
98:                     If pr > 0 Then
                            intermedio = pr
                            'o de baixo era quando era usado PVP se PVP<PR
99:                         'intermedio = System.Math.Min(intermedio, pr)
100:                    End If
                        tempcalc = Replace(a4row(17), ".", ",")
101:                    comp01_4.Text = System.Math.Min((System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)), tempcalc)
                        'comp01_4.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
102:                End If
103:
153:                End Select
154:    End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB indicar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub portaria()
1:      On Error GoTo MOSTRARERRO
2:      If codigorow(9) = True Then
3:          portimedio = "4250"
4:      End If
5:
6:      If codigorow(10) = True Then
7:          portimedio = "1234"
8:      End If
9:
10:     If codigorow(11) = True Then
11:         portimedio = "10279"
12:     End If
13:
14:     If codigorow(12) = True Then
15:         portimedio = "10280"
16:     End If
17:
18:     If codigorow(13) = True Then
19:         portimedio = "10910"
20:     End If
21:
22:     If codigorow(14) = True Then
23:         If codigorow(10) = True Then
24:             portimedio = "1234 + 14123"
25:         Else
26:             portimedio = "14123"
27:         End If
28:     End If
29:
30:     If codigorow(15) = True Then
            If port1474 = True Then
31:             portimedio = "147469"
            End If
32:     End If
33:
34:     If codigorow(19) = True Then
35:         portimedio = "21094"
36:     End If
37:
38:     If codigorow(20) = True Then
            If port1474 = True Then
39:             portimedio = "1474100"
            End If
40:     End If
41:
42:     If codigorow(21) = True Then
43:         portimedio = "6/2010"
44:     End If
45:
46:
47:     If codigorow(9) = False And codigorow(10) = False And codigorow(11) = False And codigorow(12) = False And codigorow(13) = False _
     And codigorow(14) = False And codigorow(15) = False And codigorow(19) = False And codigorow(20) = False Then
48:         portimedio = "não"
49:     End If
50:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB portaria: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Function calculo(ByVal org As Short, ByVal gen As Boolean, ByVal comp As Double, ByVal intermedio As Double) As Double
1:      On Error GoTo MOSTRARERRO
2:
3:      Select Case org
            Case 1, 2 'tipo 10
5:              calculo = intermedio * comp
6:          Case 46 'tipo 17
7:              calculo = intermedio * comp
8:          Case 42 'tipo 12
9:              calculo = intermedio
10:         Case 41 'tipo 11
11:             If comp > 0 Then
12:                 calculo = intermedio
13:             Else
14:                 calculo = 0
15:             End If
16:         Case 67 'tipo 13
17:             If comp > 0 Then
18:                 calculo = intermedio
19:             Else
20:                 calculo = 0
21:             End If
22:         Case 23, 24, 25 'diabetes - não sei se 24 e 25 também são assim mas já fica
23:             calculo = intermedio * comp
24:         Case 48, 57 'tipo 15
25:             If comp > 0 Then
26:                 If gen = "true" Then
27:                     calculo = intermedio
28:                 Else
29:                     calculo = (System.Math.Min(tectocomp, (comp + 0.15))) * intermedio
30:                 End If
31:             Else
32:                 calculo = 0
33:             End If
34:         Case 45, 59
35:             calculo = intermedio * (System.Math.Max(comp, portcomp))
36:         Case 49, 68
37:             If gen = "true" Then
38:                 calculo = intermedio
39:             Else
40:                 calculo = System.Math.Min(tectocomp, (System.Math.Max((portcomp + 0.15), (comp + 0.15)))) * intermedio
41:             End If
42:         Case 12 'SAMS
43:             calculo = intermedio * (0.9)
44:             End Select
45:
46:     'não tenho nada para o tipo 19(47) nem para os organismos 13_CGD, 15_IASFA(=01, com port=48 e com R?), 17_GNR(=01 e com R ou com port?), 18_PSP, 25_SAMSq, 05, s9(100%),
47:     '09[=(75% - 01)], CA, r1, r3, r5, 80, 81, 85, 87, 90, 95, h1[=(90% - 01)], o1, aa, ab, ae, af, xv, fm, 19
48:
49:     Exit Function
MOSTRARERRO:
        MsgBox("SUB calculo: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function




    Function SomarPVP()
1:      On Error GoTo MOSTRARERRO
2:      Dim somaPVP As Double
3:      If pvp1.Text <> "" Then
4:          pvp1v = Replace(pvp1.Text, ".", ",")
5:          pvp1val = Convert.ToDouble(pvp1v)
6:      Else : pvp1val = 0
7:      End If
8:      If pvp2.Text <> "" Then
9:          pvp2v = Replace(pvp2.Text, ".", ",")
10:         pvp2val = Convert.ToDouble(pvp2v)
11:     Else : pvp2val = 0
12:     End If
13:     If pvp3.Text <> "" Then
14:         pvp3v = Replace(pvp3.Text, ".", ",")
15:         pvp3val = Convert.ToDouble(pvp3v)
16:     Else : pvp3val = 0
17:     End If
18:     If pvp4.Text <> "" Then
19:         pvp4v = Replace(pvp4.Text, ".", ",")
20:         pvp4val = Convert.ToDouble(pvp4v)
21:     Else : pvp4val = 0
22:     End If
23:
33:     somaPVP = pvp1val + pvp2val + pvp3val + pvp4val + pvp5val + pvp6val
34:     SomarPVP = somaPVP
35:     Exit Function
MOSTRARERRO:
        MsgBox("SUB somarpvp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Function SomarComp()
1:      On Error GoTo MOSTRARERRO
2:      Dim somaComp As Double
3:      If comp01_1.Text <> "" Then
4:          comp01_1v = Replace(comp01_1.Text, ".", ",")
5:          comp01_1val = Convert.ToDouble(comp01_1v)
6:      Else : comp01_1val = 0
7:      End If
8:      If comp01_2.Text <> "" Then
9:          comp01_2v = Replace(comp01_2.Text, ".", ",")
10:         comp01_2val = Convert.ToDouble(comp01_2v)
11:     Else : comp01_2val = 0
12:     End If
13:     If comp01_3.Text <> "" Then
14:         comp01_3v = Replace(comp01_3.Text, ".", ",")
15:         comp01_3val = Convert.ToDouble(comp01_3v)
16:     Else : comp01_3val = 0
17:     End If
18:     If comp01_4.Text <> "" Then
19:         comp01_4v = Replace(comp01_4.Text, ".", ",")
20:         comp01_4val = Convert.ToDouble(comp01_4v)
21:     Else : comp01_4val = 0
22:     End If
23:
33:     somaComp = comp01_1val + comp01_2val + comp01_3val + comp01_4val + comp5val + comp6val
34:     SomarComp = somaComp
35:     Exit Function
MOSTRARERRO:
        MsgBox("SUB somarcomp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Sub somas()
1:      On Error GoTo MOSTRARERRO
2:
3:      SomarPVP()
4:      SomarComp()
5:      pvp_tot.Text = SomarPVP()
6:      comp01_tot.Text = SomarComp()
7:      If vermelho = True Then
8:          comp01_tot.BackColor = Color.Red
9:      ElseIf amarelo = True Then
10:         comp01_tot.BackColor = Color.Yellow
11:     ElseIf verde = True Then
12:         comp01_tot.BackColor = Color.Green
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB somas: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







    Private Sub AcederToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      form2.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub ECToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      EC.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub EPPRToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error GoTo MOSTRARERRO
1:      EP.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub but01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub








    Private Sub AbrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
41:
        Dim janela As New abrir
        janela.Show()
        Me.Close()
    End Sub

    Private Sub RepovoarnovosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        repovoar.Show()
    End Sub




    Sub wrapitup()
        On Error GoTo MOSTRARERRO
        Dim resultadoantes1 As String
        Dim resultadoantes2 As String
        Dim resultadoantes3 As String
        Dim resultadoantes4 As String
        Dim resultadoantes5 As String
        Dim resultadoantes6 As String
        Dim resultadodepois1 As String
        Dim resultadodepois2 As String
        Dim resultadodepois3 As String
        Dim resultadodepois4 As String
        Dim resultadodepois5 As String
        Dim resultadodepois6 As String
        Dim labelfantes1 As String
        Dim labelfantes2 As String
        Dim labelfantes3 As String
        Dim labelfantes4 As String
        Dim labelfantes5 As String
        Dim labelfantes6 As String
        Dim labelfdepois1 As String
        Dim labelfdepois2 As String
        Dim labelfdepois3 As String
        Dim labelfdepois4 As String
        Dim labelfdepois5 As String
        Dim labelfdepois6 As String
        Dim labelsubantes1 As String
        Dim labelsubantes2 As String
        Dim labelsubantes3 As String
        Dim labelsubantes4 As String
        Dim labelsubantes5 As String
        Dim labelsubantes6 As String
        Dim labelsubdepois1 As String
        Dim labelsubdepois2 As String
        Dim labelsubdepois3 As String
        Dim labelsubdepois4 As String
        Dim labelsubdepois5 As String
        Dim labelsubdepois6 As String
        Dim labellabantes1 As String
        Dim labellabantes2 As String
        Dim labellabantes3 As String
        Dim labellabantes4 As String
        Dim labellabantes5 As String
        Dim labellabantes6 As String
        Dim labellabdepois1 As String
        Dim labellabdepois2 As String
        Dim labellabdepois3 As String
        Dim labellabdepois4 As String
        Dim labellabdepois5 As String
        Dim labellabdepois6 As String
        Dim labelatribuidoantes As String
        Dim verifgenlabelantes As String
        Dim sehaportlabelantes As String
        Dim labelatribuidodepois As String
        Dim verifgenlabeldepois As String
        Dim sehaportlabeldepois As String

4:      comp01_tot.BackColor = Color.Transparent
5:
53:     amarelo = False
54:     vermelho = False
55:     verde = True
56:     Dim um As Boolean = False
57:     Dim dois As Boolean = False
58:     Dim tres As Boolean = False
59:     Dim quatro As Boolean = False
60:     Dim cinco As Boolean = False
61:     Dim seis As Boolean = False
62:     Dim sobra1 As String = ""
63:     Dim sobra2 As String = ""
64:     Dim sobra3 As String = ""
65:     'varcruza1 = 0
        'varcruza2 = 0
        'varcruza3 = 0
        'varcruza4 = 0
        'varcruza5 = 0
        'varcruza6 = 0

66:     If Not IsNothing(varcruza1) Then
67:         Select Case varcruza1
                Case Is = 1
69:                 um = True
70:             Case Is = 2
71:                 dois = True
72:             Case Is = 3
73:                 tres = True
74:             Case Is = 4
75:                 quatro = True
76:             Case Is = 5
77:                 cinco = True
78:             Case Is = 6
79:                 seis = True
80:                 End Select
81:     End If
82:     If Not IsNothing(varcruza2) Then
83:         Select Case varcruza2
                Case Is = 1
84:                 If um = False Then
85:                     um = True
86:                 Else : sobra1 = varcruza2
87:                 End If
88:             Case Is = 2
89:                 dois = True
90:             Case Is = 3
91:                 tres = True
92:             Case Is = 4
93:                 quatro = True
94:             Case Is = 5
95:                 cinco = True
96:             Case Is = 6
97:                 seis = True
98:                 End Select
99:     End If
100:    If Not IsNothing(varcruza3) Then
101:        Select Case varcruza3
                Case Is = 1
103:                If um = False Then
104:                    um = True
105:                ElseIf sobra1 = "" Then
106:                    sobra1 = varcruza3
107:                ElseIf sobra2 = "" Then
108:                    sobra2 = varcruza3
109:                Else
110:                    sobra3 = varcruza3
111:                End If
112:            Case Is = 2
113:                If dois = False Then
114:                    dois = True
115:                ElseIf sobra1 = "" Then
116:                    sobra1 = varcruza3
117:                ElseIf sobra2 = "" Then
118:                    sobra2 = varcruza3
119:                Else
120:                    sobra3 = varcruza3
121:                End If
122:            Case Is = 3
123:                tres = True
124:            Case Is = 4
125:                quatro = True
126:            Case Is = 5
127:                cinco = True
128:            Case Is = 6
129:                seis = True
130:                End Select
131:    End If
132:    If Not IsNothing(varcruza4) Then
133:        Select Case varcruza4
                Case Is = 1
135:                If um = False Then
136:                    um = True
137:                ElseIf sobra1 = "" Then
138:                    sobra1 = varcruza4
139:                ElseIf sobra2 = "" Then
140:                    sobra2 = varcruza4
141:                Else
142:                    sobra3 = varcruza4
143:                End If
144:            Case Is = 2
145:                If dois = False Then
146:                    dois = True
147:                ElseIf sobra1 = "" Then
148:                    sobra1 = varcruza4
149:                ElseIf sobra2 = "" Then
150:                    sobra2 = varcruza4
151:                Else
152:                    sobra3 = varcruza4
153:                End If
154:            Case Is = 3
155:                If tres = False Then
156:                    tres = True
157:                ElseIf sobra1 = "" Then
158:                    sobra1 = varcruza4
159:                ElseIf sobra2 = "" Then
160:                    sobra2 = varcruza4
161:                Else
162:                    sobra3 = varcruza4
163:                End If
164:            Case Is = 4
165:                quatro = True
166:            Case Is = 5
167:                cinco = True
168:            Case Is = 6
169:                seis = True
170:                End Select
171:    End If
172:    If Not IsNothing(varcruza5) Then
173:        Select Case varcruza5
                Case Is = 1
175:                If um = False Then
176:                    um = True
177:                ElseIf sobra1 = "" Then
178:                    sobra1 = varcruza5
179:                ElseIf sobra2 = "" Then
180:                    sobra2 = varcruza5
181:                Else
182:                    sobra3 = varcruza5
183:                End If
184:            Case Is = 2
185:                If dois = False Then
186:                    dois = True
187:                ElseIf sobra1 = "" Then
188:                    sobra1 = varcruza5
189:                ElseIf sobra2 = "" Then
190:                    sobra2 = varcruza5
191:                Else
192:                    sobra3 = varcruza5
193:                End If
194:            Case Is = 3
195:                If tres = False Then
196:                    tres = True
197:                ElseIf sobra1 = "" Then
198:                    sobra1 = varcruza5
199:                ElseIf sobra2 = "" Then
200:                    sobra2 = varcruza5
201:                Else
202:                    sobra3 = varcruza5
203:                End If
204:            Case Is = 4
205:                If quatro = False Then
206:                    quatro = True
207:                ElseIf sobra1 = "" Then
208:                    sobra1 = varcruza5
209:                ElseIf sobra2 = "" Then
210:                    sobra2 = varcruza5
211:                Else
212:                    sobra3 = varcruza5
213:                End If
214:            Case Is = 5
215:                cinco = True
216:            Case Is = 6
217:                seis = True
218:                End Select
219:    End If
220:    If Not IsNothing(varcruza6) Then
221:        Select Case varcruza6
                Case Is = 1
223:                If um = False Then
224:                    um = True
225:                ElseIf sobra1 = "" Then
226:                    sobra1 = varcruza6
227:                ElseIf sobra2 = "" Then
228:                    sobra2 = varcruza6
229:                Else
230:                    sobra3 = varcruza6
231:                End If
232:            Case Is = 2
233:                If dois = False Then
234:                    dois = True
235:                ElseIf sobra1 = "" Then
236:                    sobra1 = varcruza6
237:                ElseIf sobra2 = "" Then
238:                    sobra2 = varcruza6
239:                Else
240:                    sobra3 = varcruza6
241:                End If
242:            Case Is = 3
243:                If tres = False Then
244:                    tres = True
245:                ElseIf sobra1 = "" Then
246:                    sobra1 = varcruza6
247:                ElseIf sobra2 = "" Then
248:                    sobra2 = varcruza6
249:                Else
250:                    sobra3 = varcruza6
251:                End If
252:            Case Is = 4
253:                If quatro = False Then
254:                    quatro = True
255:                ElseIf sobra1 = "" Then
256:                    sobra1 = varcruza6
257:                ElseIf sobra2 = "" Then
258:                    sobra2 = varcruza6
259:                Else
260:                    sobra3 = varcruza6
261:                End If
262:            Case Is = 5
263:                If cinco = False Then
264:                    cinco = True
265:                ElseIf sobra1 = "" Then
266:                    sobra1 = varcruza6
267:                ElseIf sobra2 = "" Then
268:                    sobra2 = varcruza6
269:                Else
270:                    sobra3 = varcruza6
271:                End If
272:            Case Is = 6
273:                seis = True
274:                End Select
275:    End If
276:    If sobra1 IsNot "" Then
277:        If dois = False Then
278:            Select Case sobra1
                    Case Is = "varcruza2"
279:                    varcruza2 = 2
280:                    dois = True
281:                Case Is = "varcruza3"
282:                    varcruza3 = 2
283:                    dois = True
284:                Case Is = "varcruza2"
285:                    varcruza4 = 2
286:                    dois = True
287:                Case Is = "varcruza2"
288:                    varcruza5 = 2
289:                    dois = True
290:                Case Is = "varcruza2"
291:                    varcruza6 = 2
292:                    dois = True
293:                    End Select
294:        ElseIf tres = False Then
295:            Select Case sobra1
                    Case Is = "varcruza2"
297:                    varcruza2 = 3
298:                    tres = True
299:                Case Is = "varcruza3"
300:                    varcruza3 = 3
301:                    tres = True
302:                Case Is = "varcruza2"
303:                    varcruza4 = 3
304:                    tres = True
305:                Case Is = "varcruza2"
306:                    varcruza5 = 3
307:                    tres = True
308:                Case Is = "varcruza2"
309:                    varcruza6 = 3
310:                    tres = True
311:                    End Select
312:        ElseIf quatro = False Then
313:            Select Case sobra1
                    Case Is = "varcruza2"
315:                    varcruza2 = 4
316:                    quatro = True
317:                Case Is = "varcruza3"
318:                    varcruza3 = 4
319:                    quatro = True
320:                Case Is = "varcruza2"
321:                    varcruza4 = 4
322:                    quatro = True
323:                Case Is = "varcruza2"
324:                    varcruza5 = 4
325:                    quatro = True
326:                Case Is = "varcruza2"
327:                    varcruza6 = 4
328:                    quatro = True
                End Select
329:        ElseIf cinco = False Then
330:            Select Case sobra1
                    Case Is = "varcruza2"
332:                    varcruza2 = 5
333:                    cinco = True
334:                Case Is = "varcruza3"
335:                    varcruza3 = 5
336:                    cinco = True
337:                Case Is = "varcruza2"
338:                    varcruza4 = 5
339:                    cinco = True
340:                Case Is = "varcruza2"
341:                    varcruza5 = 5
342:                    cinco = True
343:                Case Is = "varcruza2"
344:                    varcruza6 = 5
345:                    cinco = True
346:                    End Select
347:        ElseIf seis = False Then
348:            Select Case sobra1
                    Case Is = "varcruza2"
349:                    varcruza2 = 6
350:                    seis = True
351:                Case Is = "varcruza3"
352:                    varcruza3 = 6
353:                    seis = True
354:                Case Is = "varcruza2"
355:                    varcruza4 = 6
356:                    seis = True
357:                Case Is = "varcruza2"
358:                    varcruza5 = 6
359:                    seis = True
360:                Case Is = "varcruza2"
361:                    varcruza6 = 6
362:                    seis = True
363:                    End Select
364:        End If
365:    End If
366:    If sobra2 IsNot "" Then
367:        If dois = False Then
368:            Select Case sobra2
                    Case Is = "varcruza2"
369:                    varcruza2 = 2
370:                    dois = True
371:                Case Is = "varcruza3"
372:                    varcruza3 = 2
373:                    dois = True
374:                Case Is = "varcruza2"
375:                    varcruza4 = 2
376:                    dois = True
377:                Case Is = "varcruza2"
378:                    varcruza5 = 2
379:                    dois = True
380:                Case Is = "varcruza2"
381:                    varcruza6 = 2
382:                    dois = True
383:                    End Select
384:        ElseIf tres = False Then
385:            Select Case sobra2
                    Case Is = "varcruza2"
387:                    varcruza2 = 3
388:                    tres = True
389:                Case Is = "varcruza3"
390:                    varcruza3 = 3
391:                    tres = True
392:                Case Is = "varcruza2"
393:                    varcruza4 = 3
394:                    tres = True
395:                Case Is = "varcruza2"
396:                    varcruza5 = 3
397:                    tres = True
398:                Case Is = "varcruza2"
399:                    varcruza6 = 3
400:                    tres = True
401:                    End Select
402:        ElseIf quatro = False Then
403:            Select Case sobra2
                    Case Is = "varcruza2"
405:                    varcruza2 = 4
406:                    quatro = True
407:                Case Is = "varcruza3"
408:                    varcruza3 = 4
409:                    quatro = True
410:                Case Is = "varcruza2"
411:                    varcruza4 = 4
412:                    quatro = True
413:                Case Is = "varcruza2"
414:                    varcruza5 = 4
415:                    quatro = True
416:                Case Is = "varcruza2"
417:                    varcruza6 = 4
418:                    quatro = True
419:                    End Select
420:        ElseIf cinco = False Then
421:            Select Case sobra2
                    Case Is = "varcruza2"
423:                    varcruza2 = 5
424:                    cinco = True
425:                Case Is = "varcruza3"
426:                    varcruza3 = 5
427:                    cinco = True
428:                Case Is = "varcruza2"
429:                    varcruza4 = 5
430:                    cinco = True
431:                Case Is = "varcruza2"
432:                    varcruza5 = 5
433:                    cinco = True
434:                Case Is = "varcruza2"
435:                    varcruza6 = 5
436:                    cinco = True
437:                    End Select
438:        ElseIf seis = False Then
439:            Select Case sobra2
                    Case Is = "varcruza2"
441:                    varcruza2 = 6
442:                    seis = True
443:                Case Is = "varcruza3"
444:                    varcruza3 = 6
445:                    seis = True
446:                Case Is = "varcruza2"
447:                    varcruza4 = 6
448:                    seis = True
449:                Case Is = "varcruza2"
450:                    varcruza5 = 6
451:                    seis = True
452:                Case Is = "varcruza2"
453:                    varcruza6 = 6
454:                    seis = True
455:                    End Select
456:        End If
457:    End If
458:    If sobra3 IsNot "" Then
459:        If dois = False Then
460:            Select Case sobra3
                    Case Is = "varcruza2"
461:                    varcruza2 = 2
462:                    dois = True
463:                Case Is = "varcruza3"
464:                    varcruza3 = 2
465:                    dois = True
466:                Case Is = "varcruza2"
467:                    varcruza4 = 2
468:                    dois = True
469:                Case Is = "varcruza2"
470:                    varcruza5 = 2
471:                    dois = True
472:                Case Is = "varcruza2"
473:                    varcruza6 = 2
474:                    dois = True
475:                    End Select
476:        ElseIf tres = False Then
477:            Select Case sobra3
                    Case Is = "varcruza2"
479:                    varcruza2 = 3
480:                    tres = True
481:                Case Is = "varcruza3"
482:                    varcruza3 = 3
483:                    tres = True
484:                Case Is = "varcruza2"
485:                    varcruza4 = 3
486:                    tres = True
487:                Case Is = "varcruza2"
488:                    varcruza5 = 3
489:                    tres = True
490:                Case Is = "varcruza2"
491:                    varcruza6 = 3
492:                    tres = True
493:                    End Select
494:        ElseIf quatro = False Then
495:            Select Case sobra3
                    Case Is = "varcruza2"
497:                    varcruza2 = 4
498:                    quatro = True
499:                Case Is = "varcruza3"
500:                    varcruza3 = 4
501:                    quatro = True
502:                Case Is = "varcruza2"
503:                    varcruza4 = 4
504:                    quatro = True
505:                Case Is = "varcruza2"
506:                    varcruza5 = 4
507:                    quatro = True
508:                Case Is = "varcruza2"
509:                    varcruza6 = 4
510:                    quatro = True
511:                    End Select
512:        ElseIf cinco = False Then
513:            Select Case sobra3
                    Case Is = "varcruza2"
515:                    varcruza2 = 5
516:                    cinco = True
517:                Case Is = "varcruza3"
518:                    varcruza3 = 5
519:                    cinco = True
520:                Case Is = "varcruza2"
521:                    varcruza4 = 5
522:                    cinco = True
523:                Case Is = "varcruza2"
524:                    varcruza5 = 5
525:                    cinco = True
526:                Case Is = "varcruza2"
527:                    varcruza6 = 5
528:                    cinco = True
529:                    End Select
530:        ElseIf seis = False Then
531:            Select Case sobra3
                    Case Is = "varcruza2"
533:                    varcruza2 = 6
534:                    seis = True
535:                Case Is = "varcruza3"
536:                    varcruza3 = 6
537:                    seis = True
538:                Case Is = "varcruza2"
539:                    varcruza4 = 6
540:                    seis = True
541:                Case Is = "varcruza2"
542:                    varcruza5 = 6
543:                    seis = True
544:                Case Is = "varcruza2"
545:                    varcruza6 = 6
546:                    seis = True
547:                    End Select
548:        End If
549:    End If
550:
551:
552:
553:
568:
569:
652:
653:    If vermelho = True Then
654:        comp01_tot.BackColor = Color.Red
655:    ElseIf amarelo = True Then
656:        comp01_tot.BackColor = Color.Yellow
657:    ElseIf verde = True Then
658:        comp01_tot.BackColor = Color.Green
659:    Else
660:        MsgBox("Nem vermelho, nem amarelo, nem verde")
661:    End If
662:
663:
664:    'If genverif = True Then
665:    'verifgenlabel.Text = "genéricos"
666:    'verifgenlabel.BackColor = Color.Yellow
667:    'End If
668:
669:
674:    Exit Sub
MOSTRARERRO:
        MsgBox("SUB wrapitup: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







End Class


