Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb





Public Class Form1
    Inherits Form

    Dim portcomp As Double
    Dim portcomp1 As Double
    Dim portcomp2 As Double
    Dim portcomp3 As Double
    Dim portcomp4 As Double
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
    Dim comp1v As String
    Dim comp2v As String
    Dim comp3v As String
    Dim comp4v As String
    Dim comp5v As String
    Dim comp6v As String
    Dim comp1val As Double
    Dim comp2val As Double
    Dim comp3val As Double
    Dim comp4val As Double
    Dim comp5val As Double
    Dim comp6val As Double
    Dim pr As Double
    Dim pr1 As Double
    Dim pr2 As Double
    Dim pr3 As Double
    Dim pr4 As Double
    Dim pr5 As Double
    Dim pr6 As Double
    Dim organismo As Short
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
    Dim mostracompgenports As String

    Dim codigorow As basededadosDataSet.infarmedRow
    Dim codigoarray As New ArrayList
    Dim codigo4 As New meds
    Dim novado As Boolean

    Dim form2 As New bd()

    Dim p1, p2, p3, p4, a1, a2, a3, a4 As Integer

    'trabalhar a data da comparticipação - no exemplo 40023 bate certo com 29/07/2009 - é só substituir por aXarray(5)
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



    



    Sub LimparRes()
        On Error GoTo MOSTRARERRO
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
37:     descomp1mostrado = False
38:     descomp2mostrado = False
39:     descomp3mostrado = False
40:     descomp4mostrado = False
41:     Me.aviam1.Text = ""
42:     Me.aviam2.Text = ""
43:     Me.aviam3.Text = ""
44:     Me.aviam4.Text = ""
45:     Me.result1.Text = ""
46:     Me.result2.Text = ""
47:     Me.result3.Text = ""
48:     Me.result4.Text = ""
49:     Me.aviam1.BackColor = Color.White
50:     Me.aviam2.BackColor = Color.White
51:     Me.aviam3.BackColor = Color.White
52:     Me.aviam4.BackColor = Color.White
53:     Me.result1.BackColor = Color.Transparent
54:     Me.result2.BackColor = Color.Transparent
55:     Me.result3.BackColor = Color.Transparent
56:     Me.result4.BackColor = Color.Transparent
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
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB LIMPARRES: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub LimparPrescritos()
        On Error GoTo MOSTRARERRO
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
41:     Me.result1.Text = ""
42:     Me.result2.Text = ""
43:     Me.result3.Text = ""
44:     Me.result4.Text = ""
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
65:     Me.result1.BackColor = Color.Transparent
66:     Me.result2.BackColor = Color.Transparent
67:     Me.result3.BackColor = Color.Transparent
68:     Me.result4.BackColor = Color.Transparent
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
86:     Me.semcod1.Checked = False
87:     Me.semcod1.Checked = False
88:     Me.semcod1.Checked = False
89:     Me.semcod1.Checked = False
90:     presc1gen.Checked = False
91:     presc2gen.Checked = False
92:     presc3gen.Checked = False
93:     presc4gen.Checked = False
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
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB LIMPARPRESCRITOS: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao clickar no botão limpar prescritos
    Private Sub limpresc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles limpresc.Click
        On Error GoTo MOSTRARERRO
1:      LimparPrescritos()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'o que acontece ao fazer enter no botão limpar prescritos
    Private Sub limpresc_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles limpar.Enter
        On Error GoTo MOSTRARERRO
1:      LimparPrescritos()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao fazer enter no botão limpar aviados (e resultados)
    Private Sub limpar_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles limpar.Enter
        On Error GoTo MOSTRARERRO
1:      LimparRes()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'o que acontece ao clickar no botão limpar aviados (e resultados)
    Private Sub limpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles limpar.Click
        On Error GoTo MOSTRARERRO
1:      LimparRes()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'lança a função comparar ao fazer enter no botão com o mesmo nome
    Private Sub ButComparar_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButComparar.Enter
        On Error GoTo MOSTRARERRO
1:      Comparar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'lança a função comparar ao clickar no botão com o mesmo nome
    Private Sub ButComparar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButComparar.Click
        On Error GoTo MOSTRARERRO
1:      Comparar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao abrir/'inicializar o form principal - inicia timer, limpa tudo, poe tudo a zero e foca na 1ª textbox
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo MOSTRARERRO
1:      'TODO: This line of code loads data into the 'EscolhaDataSet22.lab' table. You can move, or remove it, as needed.
2:      Me.LabTableAdapter3.Fill(Me.EscolhaDataSet22.lab)
3:      'TODO: This line of code loads data into the 'EscolhaDataSet21.lab' table. You can move, or remove it, as needed.
4:      Me.LabTableAdapter2.Fill(Me.EscolhaDataSet21.lab)
5:      'TODO: This line of code loads data into the 'EscolhaDataSet20.lab' table. You can move, or remove it, as needed.
6:      Me.LabTableAdapter.Fill(Me.EscolhaDataSet20.lab)
7:      'TODO: This line of code loads data into the 'EscolhaDataSet19.qty' table. You can move, or remove it, as needed.
8:      Me.QtyTableAdapter3.Fill(Me.EscolhaDataSet19.qty)
9:      'TODO: This line of code loads data into the 'EscolhaDataSet18.qty' table. You can move, or remove it, as needed.
10:     Me.QtyTableAdapter2.Fill(Me.EscolhaDataSet18.qty)
11:     'TODO: This line of code loads data into the 'EscolhaDataSet17.qty' table. You can move, or remove it, as needed.
12:     Me.QtyTableAdapter.Fill(Me.EscolhaDataSet17.qty)
13:     'TODO: This line of code loads data into the 'EscolhaDataSet16.forma' table. You can move, or remove it, as needed.
14:     Me.FormaTableAdapter3.Fill(Me.EscolhaDataSet16.forma)
15:     'TODO: This line of code loads data into the 'EscolhaDataSet15.forma' table. You can move, or remove it, as needed.
16:     Me.FormaTableAdapter2.Fill(Me.EscolhaDataSet15.forma)
17:     'TODO: This line of code loads data into the 'EscolhaDataSet14.forma' table. You can move, or remove it, as needed.
18:     Me.FormaTableAdapter.Fill(Me.EscolhaDataSet14.forma)
19:     'TODO: This line of code loads data into the 'EscolhaDataSet13.dc1' table. You can move, or remove it, as needed.
20:     Me.Dc1TableAdapter4.Fill(Me.EscolhaDataSet13.dc1)
21:     'TODO: This line of code loads data into the 'EscolhaDataSet12.dc1' table. You can move, or remove it, as needed.
22:     Me.Dc1TableAdapter3.Fill(Me.EscolhaDataSet12.dc1)
23:     'TODO: This line of code loads data into the 'EscolhaDataSet11.dc1' table. You can move, or remove it, as needed.
24:     Me.Dc1TableAdapter2.Fill(Me.EscolhaDataSet11.dc1)
25:     'TODO: This line of code loads data into the 'EscolhaDataSet10.dc1' table. You can move, or remove it, as needed.
26:     Me.Dc1TableAdapter1.Fill(Me.EscolhaDataSet10.dc1)
27:     'TODO: This line of code loads data into the 'EscolhaDataSet9.lab' table. You can move, or remove it, as needed.
28:     Me.LabTableAdapter1.Fill(Me.EscolhaDataSet9.lab)
29:     'TODO: This line of code loads data into the 'EscolhaDataSet8.qty' table. You can move, or remove it, as needed.
30:     Me.QtyTableAdapter1.Fill(Me.EscolhaDataSet8.qty)
31:     'TODO: This line of code loads data into the 'EscolhaDataSet7.dose' table. You can move, or remove it, as needed.
32:     Me.DoseTableAdapter1.Fill(Me.EscolhaDataSet7.dose)
33:     'TODO: This line of code loads data into the 'EscolhaDataSet6.forma' table. You can move, or remove it, as needed.
34:     Me.FormaTableAdapter1.Fill(Me.EscolhaDataSet6.forma)
35:     'TODO: This line of code loads data into the 'EscolhaDataSet5.dc1' table. You can move, or remove it, as needed.
36:     Me.Dc1TableAdapter.Fill(Me.EscolhaDataSet5.dc1)
37:     Me.InfarmedTableAdapter.Fill(Me.BasededadosDataSet.infarmed)
38:     Timer1.Start()
39:     'inicializar()
40:     Me.KeyPreview = True
41:     'Form3.Show()
42:     'Form3.SendToBack()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    'coloca o dia da semana, data e hora na label escolhida e formata a sua apresentação
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        On Error GoTo MOSTRARERRO
1:      Me.hora.Text = Format$(Now, "ddd  dd-MM-yy  HH:mm:ss")
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    'mostrar o form da bd ao clickar no botão
    Private Sub formBD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles formBD.Click
        On Error GoTo MOSTRARERRO
1:      form2.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'mostrar o form da bd ao fazer enter no botão
    Private Sub formBD_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles formBD.Enter
        On Error GoTo MOSTRARERRO
1:      form2.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao fazer enter no botão iniciar - lança 'inicializar
    Private Sub iniciar_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles iniciar.Enter
        On Error GoTo MOSTRARERRO
1:      'inicializar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'o que acontece ao clickar no botão iniciar - lança 'inicializar
    Private Sub iniciar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles iniciar.Click
        On Error GoTo MOSTRARERRO
1:      'antigamente tinha isto: Me.aviam1.Clear()
2:      'inicializar()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
8:
9:      Select Case soalgarismos
            Case False
11:             Beep()
12:             End Select
13:     Select Case foco
            Case "presc1"
15:             If semcod1.Checked = True Then
16:                 Beep()
17:                 presc1.Text = "0"
18:             End If
19:         Case "presc2"
20:             If semcod2.Checked = True Then
21:                 Beep()
22:                 presc2.Text = "0"
23:             End If
24:         Case "presc3"
25:             If semcod3.Checked = True Then
26:                 Beep()
27:                 presc3.Text = "0"
28:             End If
29:         Case "presc4"
30:             If semcod4.Checked = True Then
31:                 Beep()
32:                 presc4.Text = "0"
33:             End If
34:             End Select
35:
36:
37:
38:     If Asc(e.KeyChar) = Keys.Enter Then
39:         Select Case foco
                Case "codEC"
41:                 If codEC.Text = "" Then
42:                     Beep()
43:                     My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
44:                 Else
45:                     Me.codEC.Focus()
46:                     codEC.SelectionStart = 0
47:                     codEC.SelectionLength = Len(codEC.Text)
48:                     limpar4()
49:                     codigo4.codigo = codEC.Text
50:                     incorporar()
51:
52:                     mostrar()
53:                 End If
54:             Case "presc1"
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
73:                 Else
74:                     Me.presc3.Focus()
75:                 End If
76:             Case "presc3"
77:                 If presc3.Text = "" Then
78:                     Me.aviam1.Focus()
79:                     Me.presc3.Text = "0"
80:                     Me.presc4.Text = "0"
81:                 Else
82:                     Me.presc4.Focus()
83:                 End If
84:             Case "presc4"
85:                 If presc4.Text = "" Then
86:                     Me.aviam1.Focus()
87:                     Me.presc4.Text = "0"
88:                 Else
89:                     Me.aviam1.Focus()
90:                 End If
91:             Case "aviam2"
92:                 If aviam2.Text = "" Then
93:                     Me.aviam2.Text = "0"
94:                     Me.aviam3.Text = "0"
95:                     Me.aviam4.Text = "0"
96:                     Comparar()
97:                 Else
98:                     Me.aviam3.Focus()
99:                 End If
100:            Case "aviam3"
101:                If aviam3.Text = "" Then
102:                    Me.aviam3.Text = "0"
103:                    Me.aviam4.Text = "0"
104:                    Comparar()
105:                Else
106:                    Me.aviam4.Focus()
107:                End If
108:            Case "aviam4"
109:                If aviam4.Text = "" Then
110:                    Me.aviam4.Text = "0"
111:                End If
112:                Comparar()
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
123:        'inicializar()
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
135:        'inicializar()
136:    End If
137:
138:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.C Then
139:        Comparar()
140:    End If
141:
142:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.B Then
143:        form2.Show()
144:    End If
145:
146:    If Asc(e.KeyChar) = Keys.F11 Then
147:        If organismosRB.Checked = True Then
148:            organismosRB.Checked = False
149:            ParamiloidoseRB.Checked = True
150:        ElseIf ParamiloidoseRB.Checked = True Then
151:            ParamiloidoseRB.Checked = False
152:            organismosRB.Checked = True
153:        End If
154:    End If
155:
156:    If Asc(e.KeyChar) = Keys.Control AndAlso Asc(e.KeyChar) = Keys.O Then
157:        If organismosRB.Checked = True Then
158:            organismosRB.Checked = False
159:            ParamiloidoseRB.Checked = True
160:        ElseIf ParamiloidoseRB.Checked = True Then
161:            ParamiloidoseRB.Checked = False
162:            organismosRB.Checked = True
163:        End If
164:    End If
165:
166:    If Asc(e.KeyChar) = Keys.Escape Then
167:        If organismosRB.Checked = True Then
168:            organismosRB.Checked = False
169:            ParamiloidoseRB.Checked = True
170:        ElseIf ParamiloidoseRB.Checked = True Then
171:            ParamiloidoseRB.Checked = False
172:            organismosRB.Checked = True
173:        End If
174:    End If
175:
176:    If Asc(e.KeyChar) = Keys.Space Then
177:        'inicializar()
178:        Me.Focus()
179:        'inicializar()
180:    End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    'os próximos 8 são os validadores que estão a funcionar. fazem saltar o focus quando se inserem 7 caracteres. e lançam a comparação
    Private Sub presc1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp1 As String
2:      Caractp1 = presc1.Text
3:      If semcod1.Checked = False Then
4:          If Len(presc1.Text) = 7 Then
5:              If Caractp1 Like "#######" Then
6:                  Prescrito1.codigo = presc1.Text
7:                  'Me.presc2.Focus()
8:              End If
9:          End If
10:     Else
11:         presc1.Text = "0"
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp2 As String
2:      Caractp2 = presc2.Text
3:      If semcod2.Checked = False Then
4:          If Len(presc2.Text) = 7 Then
5:              If Caractp2 Like "#######" Then
6:                  Prescrito2.codigo = presc2.Text
7:                  'Me.presc3.Focus()
8:              End If
9:          End If
10:     Else
11:         presc2.Text = "0"
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc3.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp3 As String
2:      Caractp3 = presc3.Text
3:      If semcod3.Checked = False Then
4:          If Len(presc3.Text) = 7 Then
5:              If Caractp3 Like "#######" Then
6:                  Prescrito3.codigo = presc3.Text
7:                  'Me.presc4.Focus()
8:              End If
9:          End If
10:     Else
11:         presc3.Text = "0"
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles presc4.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractp4 As String
2:      Caractp4 = presc4.Text
3:      If semcod4.Checked = False Then
4:          If Len(presc4.Text) = 7 Then
5:              If Caractp4 Like "#######" Then
6:                  Prescrito4.codigo = presc4.Text
7:                  'Me.aviam1.Focus()
8:              End If
9:          End If
10:     Else
11:         presc4.Text = "0"
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta1 As String
2:      Caracta1 = aviam1.Text
3:      If Len(aviam1.Text) = 7 Then
4:          If Caracta1 Like "#######" Then
5:              Aviado1.codigo = aviam1.Text
6:              'Me.aviam2.Focus()
7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta2 As String
2:      Caracta2 = aviam2.Text
3:      If Len(aviam2.Text) = 7 Then
4:          If Caracta2 Like "#######" Then
5:              Aviado2.codigo = aviam2.Text
6:              'Me.aviam3.Focus()
7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam3.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta3 As String
2:      Caracta3 = aviam3.Text
3:      If Len(aviam3.Text) = 7 Then
4:          If Caracta3 Like "#######" Then
5:              Aviado3.codigo = aviam3.Text
6:              'Me.aviam4.Focus()
7:          End If
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub aviam4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles aviam4.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta4 As String
2:      Caracta4 = aviam4.Text
3:      If Len(aviam4.Text) = 7 Then
4:          If Caracta4 Like "#######" Then
5:              Aviado4.codigo = aviam4.Text
6:              'MsgBox("para compensar o enter da caneta")
7:              'Comparar()
8:          End If
9:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'chama o irbuscar() e o row2array para produzir os resultados
    Private Sub Comparar()
        On Error GoTo MOSTRARERRO
1:      Me.result1.Text = ""
2:      Me.result1.BackColor = Color.Transparent
3:      Me.result2.Text = ""
4:      Me.result2.BackColor = Color.Transparent
5:      Me.result3.Text = ""
6:      Me.result3.BackColor = Color.Transparent
7:      Me.result4.Text = ""
8:      Me.result4.BackColor = Color.Transparent
9:      irbuscar()
10:     'row2array()
11:     CorrerRegras()
12:     prioridade()
13:     AgruparAviados()
14:     AvisarDespachos()
15:     TresPresc()
16:     limparzeros()
17:     'MsgBox("a3p2.nivel = " & a3p2.nivel & vbCr & "a4p2.nivel = " & a4p2.nivel)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub limparzeros()
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
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'vai buscar à baseDeDados as linhas correspondentes aos códigos introduzidos nas 8 caixas
    Sub irbuscar()
1:      On Error GoTo MOSTRARERRO
2:      'AviadosTable = Nothing
3:      'PrescritosTable = Nothing
4:
5:
6:      If aviam4.Text <> "0" Or semcod4.Checked = True And aviam4.Text <> "" Then
7:          A = 4
8:      ElseIf aviam3.Text <> "0" Or semcod3.Checked = True And aviam3.Text <> "" Then
9:          A = 3
10:     ElseIf aviam2.Text <> "0" Or semcod2.Checked = True And aviam2.Text <> "" Then
11:         A = 2
12:     Else : A = 1
13:     End If
14:
15:     If presc4.Text <> "0" And presc4.Text <> "" Then
16:         P = 4
17:     ElseIf presc3.Text <> "0" And presc3.Text <> "" Then
18:         P = 3
19:     ElseIf presc2.Text <> "0" And presc2.Text <> "" Then
20:         P = 2
21:     Else : P = 1
22:     End If
23:
24:
25:
26:
27:     Select Case A
            Case Is = 1
29:             If semcod1.Checked = False Then
30:                 a1row = DS.infarmed.FindBycode(Aviado1.codigo)
31:                 a1array.Add(a1row)
32:                 ' Else
33:                 '    If presc1dci Then
34:                 'a1array(1) = presc1dci
35:                 'A1array(2) = presc1forma
36:                 'a1array(3) = presc1dose
37:                 'a1array(4) = presc1qty
38:                 'a1array(8) = presc1lab
40:                 'a1array(7) = presc1gen
41:                 'End If
42:             Else
43:                 'Dim file As System.IO.FileStream
44:                 'file = System.IO.File.Create("c:\ficheiro1.txt")
45:                 Dim tab As String = ","
46:                 My.Computer.FileSystem.WriteAllText _
                  ("c:\ficheiro1.txt", "9999", True)
47:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
48:                 Dim dci1 As String = presc1dci.Text
49:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", dci1, True)
50:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
51:                 Dim forma1 As String = presc1forma.Text
52:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", forma1, True)
53:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
54:                 Dim dose1 As String = presc1dose.Text
55:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", dose1, True)
56:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
57:                 Dim qty1 As String = presc1qty.Text
                    My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", qty1, True)
58:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
59:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", "99", True)
60:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
61:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", "0", True)
62:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
63:                 Dim gen1 As String = presc1gen.Text
64:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", gen1, True)
65:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", tab, True)
66:                 Dim lab1 As String = presc1lab.Text
67:                 My.Computer.FileSystem.WriteAllText _
                    ("c:\ficheiro1.txt", lab1, True)
68:                 'System.IO.File.Delete(c:\ficheiro1.txt)
69:
70:
71:                 On Error Resume Next
72:                 Dim conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\ficheiro1.mdb")
73:                 Dim cmd As New OleDbCommand("SELECT * INTO [um] FROM [Text;Database=c:\;Hdr=No].[ficheiro1.txt]", conn)
74:                 conn.Open()
75:                 cmd.ExecuteNonQuery()
76:                 conn.Close()
77:                 On Error GoTo MOSTRARERRO
78:                 'Me.umTableAdapter.Fill(Me.ficheiro1DataSet.um)
79:                 'a1row = FS.um.FindBycode("ï»¿9999")
80:                 a1array.Add(a1row)
81:
82:
83:                 '    If presc1dci Then
84:                 'a1array(1) = presc1dci
85:                 'A1array(2) = presc1forma
86:                 'a1array(3) = presc1dose
87:                 'a1array(4) = presc1qty
88:                 'a1array(8) = presc1lab
89:                 'a1array(7) = presc1gen
90:                 'End If
91:
92:
93:                 'a1array.Add(linha1)
94:             End If
95:         Case Is = 2
96:
97:             If semcod1.Checked = False Then
98:                 a1row = DS.infarmed.FindBycode(Aviado1.codigo)
99:                 a1array.Add(a1row)
100:            Else
101:
102:                'a1array.Add(linha1)
103:            End If
104:            If semcod2.Checked = False Then
105:                a2row = DS.infarmed.FindBycode(Aviado2.codigo)
106:                a2array.Add(a2row)
107:            Else
108:                'a2array.Add(linha2)
109:            End If
110:        Case Is = 3
111:
112:            If semcod1.Checked = False Then
113:                a1row = DS.infarmed.FindBycode(Aviado1.codigo)
114:                a1array.Add(a1row)
115:            Else
116:
117:                'a1array.Add(linha1)
118:            End If
119:            If semcod2.Checked = False Then
120:                a2row = DS.infarmed.FindBycode(Aviado2.codigo)
121:                a2array.Add(a2row)
122:            Else
123:                'a2array.Add(linha2)
124:            End If
125:            If semcod3.Checked = False Then
126:                a3row = DS.infarmed.FindBycode(Aviado3.codigo)
127:                a3array.Add(a3row)
128:            Else
129:                'a3array.Add(linha3)
130:            End If
131:        Case Is = 4
132:
133:            If semcod1.Checked = False Then
134:                a1row = DS.infarmed.FindBycode(Aviado1.codigo)
135:                a1array.Add(a1row)
136:            Else
137:                'a1array.Add(linha1)
138:            End If
139:            If semcod2.Checked = False Then
140:                a2row = DS.infarmed.FindBycode(Aviado2.codigo)
141:                a2array.Add(a2row)
142:            Else
143:                'a2array.Add(linha2)
144:            End If
145:            If semcod3.Checked = False Then
146:                a3row = DS.infarmed.FindBycode(Aviado3.codigo)
147:                a3array.Add(a3row)
148:            Else
149:                ' a3array.Add(linha3)
150:            End If
151:            If semcod4.Checked = False Then
152:                a4row = DS.infarmed.FindBycode(Aviado4.codigo)
153:                a4array.Add(a4row)
154:            Else
155:                'a4array.Add(linha4)
156:            End If
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub CorrerRegras()
1:      On Error GoTo MOSTRARERRO
2:      Select Case P 'ver quantos foram prescritos
            Case Is = 1 'se só foi prescrita uma embalagem
4:              If Not IsNothing(p1row) Then
5:                  Dim p1array = p1row.ItemArray
6:                  p1a1234()
7:              Else
8:                  prescNexist(1)
9:              End If
10:         Case Else 'caso tenham sido prescritos mais de uma embalagem
11:             If P > 1 Then
12:                 If Not IsNothing(p1row) Then
13:                     Dim p1array = p1row.ItemArray
14:                     If A >= 1 Then
15:                         If Not IsNothing(a1row) Then 'Se existir o código do 1º aviado
16:                             Dim a1array = a1row.ItemArray
17:                             If a1array(5).ToString = 0 And paramiloidose = False And a1array(9) = False Then 'se o 1º aviado não for comparticipado
18:                                 nComp(1)
19:                             Else
20:                                 p1a1()
21:                                 If a1array(5).ToString <> 0 And a1array(5).ToString <> 15 And a1array(5).ToString <> 37 And a1array(5).ToString <> 69 _
         And a1array(5).ToString <> 95 And a1array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
23:                                     aviam1.BackColor = Color.Yellow
24:                                     If descomp1mostrado = False Then
25:                                         MsgBox(a1row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a1array(5), #12/30/1899#) & ".")
26:                                         descomp1mostrado = True
27:                                     End If
28:                                 End If
29:                             End If
30:
31:                         Else
32:                             aviadoNexist(1)
33:                             a1p1.nivel = 0
34:                             a1p2.nivel = 0
35:                             a1p3.nivel = 0
36:                             a1p4.nivel = 0
37:                             a1p1.resultado = 10
38:                             a1p2.resultado = 10
39:                             a1p3.resultado = 10
40:                             a1p4.resultado = 10
41:                         End If
42:                     End If
43:
44:                     If A >= 2 Then
45:                         If Not IsNothing(a2row) Then 'Se existir o código do 2º aviado
46:                             Dim a2array = a2row.ItemArray
47:                             If a2array(5).ToString = 0 And paramiloidose = False And a2array(9) = False Then 'se o 2º aviado não for comparticipado
48:                                 nComp(2)
49:                             Else
50:                                 p1a2()
51:                                 If a2array(5).ToString <> 0 And a2array(5).ToString <> 15 And a2array(5).ToString <> 37 And a2array(5).ToString <> 69 _
   And a2array(5).ToString <> 95 And a2array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
53:                                     aviam2.BackColor = Color.Yellow
54:                                     If descomp2mostrado = False Then
55:                                         MsgBox(a2row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a2array(5), #12/30/1899#) & ".")
56:                                         descomp2mostrado = True
57:                                     End If
58:                                 End If
59:                             End If
60:                         Else
61:                             aviadoNexist(2)
62:                             a2p1.nivel = 0
63:                             a2p2.nivel = 0
64:                             a2p3.nivel = 0
65:                             a2p4.nivel = 0
66:                             a2p1.resultado = 10
67:                             a2p2.resultado = 10
68:                             a2p3.resultado = 10
69:                             a2p4.resultado = 10
70:                         End If
71:                     End If
72:                     If A >= 3 Then
73:                         If Not IsNothing(a3row) Then 'Se existir o código do 3º aviado
74:                             Dim a3array = a3row.ItemArray
75:                             If a3array(5).ToString = 0 And paramiloidose = False And a3array(9) = False Then 'se o 3º aviado não for comparticipado
76:                                 nComp(3)
77:                             Else
78:                                 p1a3()
79:                                 If a3array(5).ToString <> 0 And a3array(5).ToString <> 15 And a3array(5).ToString <> 37 And a3array(5).ToString <> 69 _
             And a3array(5).ToString <> 95 And a3array(5).ToString <> 100 Then 'se a comp do 3º aviado for destas %
80:                                     aviam3.BackColor = Color.Yellow
81:                                     If descomp3mostrado = False Then
82:                                         MsgBox(a3row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a3array(5), #12/30/1899#) & ".")
83:                                         descomp3mostrado = True
84:                                     End If
85:                                 End If
86:
87:                             End If
88:                         Else
89:                             aviadoNexist(3)
90:                             a3p1.nivel = 0
91:                             a3p2.nivel = 0
92:                             a3p3.nivel = 0
93:                             a3p4.nivel = 0
94:                             a3p1.resultado = 10
95:                             a3p2.resultado = 10
96:                             a3p3.resultado = 10
97:                             a3p4.resultado = 10
98:                         End If
99:                     End If
100:
101:                    If A >= 4 Then
102:                        If Not IsNothing(a4row) Then 'Se existir o código do 4º aviado
103:                            Dim a4array = a4row.ItemArray
104:                            If a4array(5).ToString = 0 And paramiloidose = False And a4array(9) = False Then 'se o 4º aviado não for comparticipado
105:                                nComp(4)
106:                            Else
107:                                p1a4()
108:                                If a4array(5).ToString <> 0 And a4array(5).ToString <> 15 And a4array(5).ToString <> 37 And a4array(5).ToString <> 69 _
             And a4array(5).ToString <> 95 And a4array(5).ToString <> 100 Then 'se a comp do 4º aviado for destas %
109:                                    aviam4.BackColor = Color.Yellow
110:                                    If descomp4mostrado = False Then
111:                                        MsgBox(a4row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a4array(5), #12/30/1899#) & ".")
112:                                        descomp4mostrado = True
113:                                    End If
114:                                End If
115:                            End If
116:                        Else
117:                            aviadoNexist(4)
118:                            a4p1.nivel = 0
119:                            a4p2.nivel = 0
120:                            a4p3.nivel = 0
121:                            a4p4.nivel = 0
122:                            a4p1.resultado = 10
123:                            a4p2.resultado = 10
124:                            a4p3.resultado = 10
125:                            a4p4.resultado = 10
126:                        End If
127:                    End If
128:                End If
129:            End If
130:            If P >= 2 Then
131:                If Not IsNothing(p2row) Then
132:                    Dim p2array = p2row.ItemArray
133:                    If A >= 1 Then
134:                        If Not IsNothing(a1row) Then 'Se existir o código do 1º aviado
135:                            Dim a1array = a1row.ItemArray
136:                            If a1array(5).ToString = 0 And paramiloidose = False And a1array(9) = False Then 'se o 1º aviado não for comparticipado
137:                                nComp(1)
138:                            Else
139:                                p2a1()
140:                                If a1array(5).ToString <> 0 And a1array(5).ToString <> 15 And a1array(5).ToString <> 37 And a1array(5).ToString <> 69 _
             And a1array(5).ToString <> 95 And a1array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
142:                                    aviam1.BackColor = Color.Yellow
143:                                    If descomp1mostrado = False Then
144:                                        MsgBox(a1row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a1array(5), #12/30/1899#) & ".")
145:                                        descomp1mostrado = True
146:                                    End If
147:                                End If
148:                            End If
149:                        Else
150:                            aviadoNexist(1)
151:                            a1p1.nivel = 0
152:                            a1p2.nivel = 0
153:                            a1p3.nivel = 0
154:                            a1p4.nivel = 0
155:                            a1p1.resultado = 10
156:                            a1p2.resultado = 10
157:                            a1p3.resultado = 10
158:                            a1p4.resultado = 10
159:                        End If
160:                    End If
161:
162:                    If A >= 2 Then
163:                        If Not IsNothing(a2row) Then 'Se existir o código do 2º aviado
164:                            Dim a2array = a2row.ItemArray
165:                            If a2array(5).ToString = 0 And paramiloidose = False And a2array(9) = False Then 'se o 2º aviado não for comparticipado
166:                                nComp(2)
167:                            Else
168:                                p2a2()
169:                                If a2array(5).ToString <> 0 And a2array(5).ToString <> 15 And a2array(5).ToString <> 37 And a2array(5).ToString <> 69 _
             And a2array(5).ToString <> 95 And a2array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
171:                                    aviam2.BackColor = Color.Yellow
172:                                    If descomp2mostrado = False Then
173:                                        MsgBox(a2row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a2array(5), #12/30/1899#) & ".")
174:                                        descomp2mostrado = True
175:                                    End If
176:                                End If
177:                            End If
178:                        Else
179:                            aviadoNexist(2)
180:                            a2p1.nivel = 0
181:                            a2p2.nivel = 0
182:                            a2p3.nivel = 0
183:                            a2p4.nivel = 0
184:                            a2p1.resultado = 10
185:                            a2p2.resultado = 10
186:                            a2p3.resultado = 10
187:                            a2p4.resultado = 10
188:                        End If
189:                    End If
190:
191:                    If A >= 3 Then
192:                        If Not IsNothing(a3row) Then 'Se existir o código do 3º aviado
193:                            Dim a3array = a3row.ItemArray
194:                            If a3array(5).ToString = 0 And paramiloidose = False And a3array(9) = False Then 'se o 3º aviado não for comparticipado
195:                                nComp(3)
196:                            Else
197:                                p2a3()
198:                                If a3array(5).ToString <> 0 And a3array(5).ToString <> 15 And a3array(5).ToString <> 37 And a3array(5).ToString <> 69 _
           And a3array(5).ToString <> 95 And a3array(5).ToString <> 100 Then 'se a comp do 3º aviado for destas %
200:                                    aviam3.BackColor = Color.Yellow
201:                                    If descomp3mostrado = False Then
202:                                        MsgBox(a3row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a3array(5), #12/30/1899#) & ".")
203:                                        descomp3mostrado = True
204:                                    End If
205:                                End If
206:                            End If
207:                        Else
208:                            aviadoNexist(3)
209:                            a3p1.nivel = 0
210:                            a3p2.nivel = 0
211:                            a3p3.nivel = 0
212:                            a3p4.nivel = 0
213:                            a3p1.resultado = 10
214:                            a3p2.resultado = 10
215:                            a3p3.resultado = 10
216:                            a3p4.resultado = 10
217:                        End If
218:                    End If
219:
220:                    If A >= 4 Then
221:                        If Not IsNothing(a4row) Then 'Se existir o código do 4º aviado
222:                            Dim a4array = a4row.ItemArray
223:                            If a4array(5).ToString = 0 And paramiloidose = False And a4array(9) = False Then 'se o 4º aviado não for comparticipado
224:                                nComp(4)
225:                            Else
226:                                p2a4()
227:                                If a4array(5).ToString <> 0 And a4array(5).ToString <> 15 And a4array(5).ToString <> 37 And a4array(5).ToString <> 69 _
             And a4array(5).ToString <> 95 And a4array(5).ToString <> 100 Then 'se a comp do 4º aviado for destas %
228:                                    aviam4.BackColor = Color.Yellow
229:                                    If descomp4mostrado = False Then
230:                                        MsgBox(a4row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a4array(5), #12/30/1899#) & ".")
231:                                        descomp4mostrado = True
232:                                    End If
233:                                End If
234:                            End If
235:                        Else
236:                            aviadoNexist(4)
237:                            a4p1.nivel = 0
238:                            a4p2.nivel = 0
239:                            a4p3.nivel = 0
240:                            a4p4.nivel = 0
241:                            a4p1.resultado = 10
242:                            a4p2.resultado = 10
243:                            a4p3.resultado = 10
244:                            a4p4.resultado = 10
245:                        End If
246:                    End If
247:                End If
248:            End If
249:
250:
251:
252:            If P >= 3 Then
253:                If Not IsNothing(p3row) Then
254:                    Dim p3array = p3row.ItemArray
255:                    If A >= 1 Then
256:                        If Not IsNothing(a1row) Then 'Se existir o código do 1º aviado
257:                            Dim a1array = a1row.ItemArray
258:                            If a1array(5).ToString = 0 And paramiloidose = False And a1array(9) = False Then 'se o 1º aviado não for comparticipado
259:                                nComp(1)
260:                            Else
261:                                p3a1()
262:                                If a1array(5).ToString <> 0 And a1array(5).ToString <> 15 And a1array(5).ToString <> 37 And a1array(5).ToString <> 69 _
         And a1array(5).ToString <> 95 And a1array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
264:                                    aviam1.BackColor = Color.Yellow
265:                                    If descomp1mostrado = False Then
266:                                        MsgBox(a1row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a1array(5), #12/30/1899#) & ".")
267:                                        descomp1mostrado = True
268:                                    End If
269:                                End If
270:                            End If
271:                        Else
272:                            aviadoNexist(1)
273:                            a1p1.nivel = 0
274:                            a1p2.nivel = 0
275:                            a1p3.nivel = 0
276:                            a1p4.nivel = 0
277:                            a1p1.resultado = 10
278:                            a1p2.resultado = 10
279:                            a1p3.resultado = 10
280:                            a1p4.resultado = 10
281:                        End If
282:                    End If
283:
284:                    If A >= 2 Then
285:                        If Not IsNothing(a2row) Then 'Se existir o código do 2º aviado
286:                            Dim a2array = a2row.ItemArray
287:                            If a2array(5).ToString = 0 And paramiloidose = False And a2array(9) = False Then 'se o 2º aviado não for comparticipado
288:                                nComp(2)
289:                            Else
290:                                p3a2()
291:                                If a2array(5).ToString <> 0 And a2array(5).ToString <> 15 And a2array(5).ToString <> 37 And a2array(5).ToString <> 69 _
             And a2array(5).ToString <> 95 And a2array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
292:                                    aviam2.BackColor = Color.Yellow
293:                                    If descomp2mostrado = False Then
294:                                        MsgBox(a2row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a2array(5), #12/30/1899#) & ".")
295:                                        descomp2mostrado = True
296:                                    End If
297:                                End If
298:                            End If
299:                        Else
300:                            aviadoNexist(2)
301:                            a2p1.nivel = 0
302:                            a2p2.nivel = 0
303:                            a2p3.nivel = 0
304:                            a2p4.nivel = 0
305:                            a2p1.resultado = 10
306:                            a2p2.resultado = 10
307:                            a2p3.resultado = 10
308:                            a2p4.resultado = 10
309:                        End If
310:                    End If
311:
312:                    If A >= 3 Then
313:                        If Not IsNothing(a3row) Then 'Se existir o código do 3º aviado
314:                            Dim a3array = a3row.ItemArray
315:                            If a3array(5).ToString = 0 And paramiloidose = False And a3array(9) = False Then 'se o 3º aviado não for comparticipado
316:                                nComp(3)
317:                            Else
318:                                p3a3()
319:                                If a3array(5).ToString <> 0 And a3array(5).ToString <> 15 And a3array(5).ToString <> 37 And a3array(5).ToString <> 69 _
             And a3array(5).ToString <> 95 And a3array(5).ToString <> 100 Then 'se a comp do 3º aviado for destas %
320:                                    aviam3.BackColor = Color.Yellow
321:                                    If descomp3mostrado = False Then
322:                                        MsgBox(a3row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a3array(5), #12/30/1899#) & ".")
323:                                        descomp3mostrado = True
324:                                    End If
325:                                End If
326:                            End If
327:                        Else
328:                            aviadoNexist(3)
329:                            a3p1.nivel = 0
330:                            a3p2.nivel = 0
331:                            a3p3.nivel = 0
332:                            a3p4.nivel = 0
333:                            a3p1.resultado = 10
334:                            a3p2.resultado = 10
335:                            a3p3.resultado = 10
336:                            a3p4.resultado = 10
337:                        End If
338:                    End If
339:
340:                    If A >= 4 Then
341:                        If Not IsNothing(a4row) Then 'Se existir o código do 4º aviado
342:                            Dim a4array = a4row.ItemArray
343:                            If a4array(5).ToString = 0 And paramiloidose = False And a4array(9) = False Then 'se o 4º aviado não for comparticipado
344:                                nComp(4)
345:                            Else
346:                                p3a4()
347:                                If a4array(5).ToString <> 0 And a4array(5).ToString <> 15 And a4array(5).ToString <> 37 And a4array(5).ToString <> 69 _
             And a4array(5).ToString <> 95 And a4array(5).ToString <> 100 Then 'se a comp do 4º aviado for destas %
348:                                    aviam4.BackColor = Color.Yellow
349:                                    If descomp4mostrado = False Then
350:                                        MsgBox(a4row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a4array(5), #12/30/1899#) & ".")
351:                                        descomp4mostrado = True
352:                                    End If
353:                                End If
354:                            End If
355:                        Else
356:                            aviadoNexist(4)
357:                            a4p1.nivel = 0
358:                            a4p2.nivel = 0
359:                            a4p3.nivel = 0
360:                            a4p4.nivel = 0
361:                            a4p1.resultado = 10
362:                            a4p2.resultado = 10
363:                            a4p3.resultado = 10
364:                            a4p4.resultado = 10
365:                        End If
366:                    End If
367:                End If
368:            End If
369:
370:
371:            If P >= 4 Then
372:                If Not IsNothing(p4row) Then
373:                    Dim p4array = p4row.ItemArray
374:                    If A >= 1 Then
375:                        If Not IsNothing(a1row) Then 'Se existir o código do 1º aviado
376:                            Dim a1array = a1row.ItemArray
377:                            If a1array(5).ToString = 0 And paramiloidose = False And a1array(9) = False Then 'se o 1º aviado não for comparticipado
378:                                nComp(1)
379:                            Else
380:                                p4a1()
381:                                If a1array(5).ToString <> 0 And a1array(5).ToString <> 15 And a1array(5).ToString <> 37 And a1array(5).ToString <> 69 _
             And a1array(5).ToString <> 95 And a1array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
382:                                    aviam1.BackColor = Color.Yellow
383:                                    If descomp1mostrado = False Then
384:                                        MsgBox(a1row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a1array(5), #12/30/1899#) & ".")
385:                                        descomp1mostrado = True
386:                                    End If
387:                                End If
388:                            End If
389:                        Else
390:                            aviadoNexist(1)
391:                            a1p1.nivel = 0
392:                            a1p2.nivel = 0
393:                            a1p3.nivel = 0
394:                            a1p4.nivel = 0
395:                            a1p1.resultado = 10
396:                            a1p2.resultado = 10
397:                            a1p3.resultado = 10
398:                            a1p4.resultado = 10
399:                        End If
400:                    End If
401:
402:                    If A >= 2 Then
403:                        If Not IsNothing(a2row) Then 'Se existir o código do 2º aviado
404:                            Dim a2array = a2row.ItemArray
405:                            If a2array(5).ToString = 0 And paramiloidose = False And a2array(9) = False Then 'se o 2º aviado não for comparticipado
406:                                nComp(2)
407:                            Else
408:                                p4a2()
409:                                If a2array(5).ToString <> 0 And a2array(5).ToString <> 15 And a2array(5).ToString <> 37 And a2array(5).ToString <> 69 _
             And a2array(5).ToString <> 95 And a2array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
410:                                    aviam2.BackColor = Color.Yellow
411:                                    If descomp2mostrado = False Then
412:                                        MsgBox(a2row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a2array(5), #12/30/1899#) & ".")
413:                                        descomp2mostrado = True
414:                                    End If
415:                                End If
416:                            End If
417:                        Else
418:                            aviadoNexist(2)
419:                            a2p1.nivel = 0
420:                            a2p2.nivel = 0
421:                            a2p3.nivel = 0
422:                            a2p4.nivel = 0
423:                            a2p1.resultado = 10
424:                            a2p2.resultado = 10
425:                            a2p3.resultado = 10
426:                            a2p4.resultado = 10
427:                        End If
428:                    End If
429:
430:                    If A >= 3 Then
431:                        If Not IsNothing(a3row) Then 'Se existir o código do 3º aviado
432:                            Dim a3array = a3row.ItemArray
433:                            If a3array(5).ToString = 0 And paramiloidose = False And a3array(9) = False Then 'se o 3º aviado não for comparticipado
434:                                nComp(3)
435:                            Else
436:                                p4a3()
437:                                If a3array(5).ToString <> 0 And a3array(5).ToString <> 15 And a3array(5).ToString <> 37 And a3array(5).ToString <> 69 _
             And a3array(5).ToString <> 95 And a3array(5).ToString <> 100 Then 'se a comp do 3º aviado for destas %
438:                                    aviam3.BackColor = Color.Yellow
439:                                    If descomp3mostrado = False Then
440:                                        MsgBox(a3row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a3array(5), #12/30/1899#) & ".")
441:                                        descomp3mostrado = True
442:                                    End If
443:                                End If
444:                            End If
445:                        Else
446:                            aviadoNexist(3)
447:                            a3p1.nivel = 0
448:                            a3p2.nivel = 0
449:                            a3p3.nivel = 0
450:                            a3p4.nivel = 0
451:                            a3p1.resultado = 10
452:                            a3p2.resultado = 10
453:                            a3p3.resultado = 10
454:                            a3p4.resultado = 10
455:                        End If
456:                    End If
457:
458:                    If A >= 4 Then
459:                        If Not IsNothing(a4row) Then 'Se existir o código do 4º aviado
460:                            Dim a4array = a4row.ItemArray
461:                            If a4array(5).ToString = 0 And paramiloidose = False And a4array(9) = False Then 'se o 4º aviado não for comparticipado
462:                                nComp(4)
463:                            Else
464:                                p4a4()
465:                                If a4array(5).ToString <> 0 And a4array(5).ToString <> 15 And a4array(5).ToString <> 37 And a4array(5).ToString <> 69 _
             And a4array(5).ToString <> 95 And a4array(5).ToString <> 100 Then 'se a comp do 4º aviado for destas %
466:                                    aviam4.BackColor = Color.Yellow
467:                                    If descomp4mostrado = False Then
468:                                        MsgBox(a4row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a4array(5), #12/30/1899#) & ".")
469:                                        descomp4mostrado = True
470:                                    End If
471:                                End If
472:                            End If
473:                        Else
474:                            aviadoNexist(4)
475:                            a4p1.nivel = 0
476:                            a4p2.nivel = 0
477:                            a4p3.nivel = 0
478:                            a4p4.nivel = 0
479:                            a4p1.resultado = 10
480:                            a4p2.resultado = 10
481:                            a4p3.resultado = 10
482:                            a4p4.resultado = 10
483:                        End If
484:                    End If
485:                End If
486:            End If
487:
488:            End Select
489:    Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







    Sub p1a1234()
1:      On Error GoTo MOSTRARERRO
2:      If Not IsNothing(p1row) Then 'se o código do 1º prescrito existir
3:          Dim p1array = p1row.ItemArray
4:
5:          If A = 1 Then 'se só foi aviada uma embalagem
6:
7:              If Not IsNothing(a1row) Then 'se o código do 1º aviado existir
8:                  Dim a1array = a1row.ItemArray
9:                  If a1array(5).ToString = 0 And paramiloidose = False And a1array(9) = False Then 'se o 1º aviado não for comparticipado
10:                     nComp(1)
11:                 Else 'se o primeiro aviado for comparticipado
12:                     If a1array(5).ToString <> 0 And a1array(5).ToString <> 15 And a1array(5).ToString <> 37 And a1array(5).ToString <> 69 _
    And a1array(5).ToString <> 95 And a1array(5).ToString <> 100 Then 'se a comp do 1º aviado for destas %
13:                         aviam1.BackColor = Color.Yellow
14:                         If descomp1mostrado = False Then
15:                             MsgBox(a1row(0) & " em descomparticipação / escoamento até " & DateAdd(DateInterval.Day, a1array(5), #12/30/1899#) & ".")
16:                             descomp1mostrado = True
17:                         End If
18:                     End If
19:                     'If a1array(10) = True Or a1array(11) = True Or a1array(12) = True Or a1array(13) = True Or a1array(14) = True Or a1array(15) = True Then
20:                     'If a1array(9) = True Then
21:                     '  alzheimer(a1array(0))
22:                     ' End If
23:                     'If a1array(10) = True Then
24:                     'gastro(a1array(0))
25:                     '    End If
26:                     'If a1array(11) = True Then
27:                     'espondilite(a1array(0))
28:                     '  End If
29:                     'If a1array(15) = True Then
30:                     'espondilite(a1array(0))
31:                     'End If
32:                     'If a1array(12) = True Then
33:                     ' despacho(a1array(0), "10279/2008, de 11/03")
34:                     'End If
35:                     'If a1array(13) = True Then
36:                     'despacho(a1array(0), "10280/2008, de 11/03")
37:                     'End If
38:                     'If a1array(14) = True Then
39:                     'despacho(a1array(0), "10910/2009, de 22/04")
40:                     'End If
41:                     'End If
42:                     Select Case a1array(0).ToString 'ver o código do 1º aviado 
                            Case p1array(0).ToString  'no caso do código do 1º aviado ser igual ao do 1º prescrito
44:                             If procurado1 = True Then
45:                                 OK(1)
46:                                 anular(1)
47:                                 av1.nivel = 0
48:                             End If
49:                         Case Else 'no caso do código do 1º aviado NÃO ser igual ao do 1º prescrito
50:                             Select Case a1array(1).ToString 'ver o dci do 1º aviado
                                    Case p1array(1).ToString 'no caso do dci do 1º aviado ser igual ao do 1º prescrito
52:                                     If procurado1 Then
53:                                         If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString And _
                                     a1array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 1º aviado forem todos iguais ao do 1º prescrito
55:                                             Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
57:                                                     OK(1)
58:                                                     anular(1)
59:                                                     av1.nivel = 0
60:                                                 Case 2
61:                                                     marca(1)
62:                                                     av1.nivel = 2
63:                                                 Case 3
64:                                                     marca2gen(1)
65:                                                     av1.nivel = 2
66:                                                 Case 4
67:                                                     result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
68:                                                     result1.BackColor = Color.Yellow
69:                                                     av1.nivel = 2
70:                                                     End Select
71:                                         Else 'se forma, dose e qty do 1º aviado NÃO forem todos iguais ao do 1º prescrito
72:                                             If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString Then
73:                                                 'se forma e dose do 1º aviado forem todos iguais ao do 1º prescrito
74:                                                 If a1array(4).ToString = p1array(4).ToString Then
75:                                                     Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                            Case 1
77:                                                             OK(1)
78:                                                             anular(1)
79:                                                             av1.nivel = 0
80:                                                         Case 2
81:                                                             marca(1)
82:                                                             av1.nivel = 2
83:                                                         Case 3
84:                                                             marca2gen(1)
85:                                                             av1.nivel = 2
86:                                                         Case 4
87:                                                             result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
88:                                                             result1.BackColor = Color.Yellow
89:                                                             av1.nivel = 2
90:                                                             End Select
91:                                                 ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p1array(4).ToString Then 'se a qty do 1º aviado contiver apenas algarismos
92:                                                     If a1array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 1º aviado for <= 150% a do 1º prescrito
93:                                                         Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                                Case 1
95:                                                                 OK(1)
96:                                                                 anular(1)
97:                                                                 av1.nivel = 0
98:                                                             Case 2
99:                                                                 marca(1)
100:                                                                av1.nivel = 2
101:                                                            Case 3
102:                                                                marca2gen(1)
103:                                                                av1.nivel = 2
104:                                                            Case 4
105:                                                                result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
106:                                                                result1.BackColor = Color.Yellow
107:                                                                av1.nivel = 0
108:                                                                End Select
110:                                                    Else 'se a qty do 1º aviado for > 150% a do 1º prescrito
111:                                                        Hquant(1)
112:                                                        av1.nivel = 3
113:                                                    End If
114:                                                Else 'se a qty do 1º aviado NÃO contiver apenas algarismos
115:                                                    verifQuant(1)
116:                                                    av1.nivel = 3
117:                                                End If
118:                                            Else  'se forma e dose do 1º aviado NÃO forem todos iguais ao do 1º prescrito
119:                                                If Via(a1array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 1º aviado for diferente da do 1º prescrito
120:                                                    apresDif(1)
121:                                                Else
122:                                                    If a1array(3).ToString <> p1array(3).ToString Then 'se dose do 1º aviado for diferente da do 1º prescrito
123:                                                        DoseDif(1)
124:                                                    End If
125:                                                End If
126:                                            End If
127:                                        End If
128:                                    End If
129:                                Case Else 'no caso do dci do 1º aviado NÃO ser igual ao do 1º prescrito
130:                                    dciDif(1)
131:                                    End Select
132:                            End Select
133:                End If
134:            Else 'se o código do 1º aviado NÂO existir
135:                aviadoNexist(1)
136:            End If
137:        End If
138:        If A = 2 Then
139:            If Not IsNothing(a1row) And Not IsNothing(a2row) Then 'se o código de ambos os aviados existe
140:                Dim a1array = a1row.ItemArray
141:                Dim a2array = a2row.ItemArray
142:                Select Case p1array(0) 'ver qual o código do 1º prescrito
                        Case a1array(0) 'no caso do código do 1º prescrito ser igual ao do 1º aviado
143:                        If procurado1 = True Then
144:                            OK(1)
145:                            anular(1)
146:                            av1.nivel = 0
147:                        End If
148:                    Case a2array(0) 'no caso do código do 1º prescrito ser igual ao do 2º aviado
149:                        If procurado1 = True Then
150:                            OK(2)
151:                            anular(1)
152:                            av2.nivel = 0
153:                        End If
154:                    Case Else 'no caso do código do 1º prescrito NÃO ser igual a nenhum dos dois aviados
155:                        'If procurado1 = True Then
156:
157:                        'Select Case p1array(1) 'ver qual o dci do 1º prescrito
158:                        'Case a1array(1) 'no caso do dci do 1º prescrito ser igual ao do 1º aviado
159:                        If p1array(1) = a1array(1) Then
160:                            If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString And _
                             a1array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 1º aviado forem todos iguais ao do 1º prescrito
162:                                Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
164:                                        OK(1)
165:                                        anular(1)
166:                                        av1.nivel = 0
167:                                    Case 2
168:                                        marca(1)
169:                                        av1.nivel = 2
170:                                    Case 3
171:                                        marca2gen(1)
172:                                        av1.nivel = 2
173:                                    Case 4
174:                                        result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
175:                                        result1.BackColor = Color.Yellow
176:                                        av1.nivel = 2
177:                                        End Select
178:                            Else 'se forma, dose e qty do 1º aviado NÃO forem todos iguais ao do 1º prescrito
179:                                If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString Then
180:                                    'se forma e dose do 1º aviado forem todos iguais ao do 1º prescrito
181:                                    If a1array(4).ToString = p1array(4).ToString Then
182:                                        Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
184:                                                OK(1)
185:                                                anular(1)
186:                                                av1.nivel = 0
187:                                            Case 2
188:                                                marca(1)
189:                                                av1.nivel = 2
190:                                            Case 3
191:                                                marca2gen(1)
192:                                                av1.nivel = 2
193:                                            Case 4
194:                                                result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
195:                                                result1.BackColor = Color.Yellow
196:                                                av1.nivel = 2
197:                                                End Select
198:                                    ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p1array(4).ToString Then 'se a qty do 1º aviado contiver apenas algarismos
199:                                        If a1array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 1º aviado for <= 150% a do 1º prescrito
200:                                            Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
202:                                                    OK(1)
203:                                                    anular(1)
204:                                                    av1.nivel = 0
205:                                                Case 2
206:                                                    marca(1)
207:                                                    av1.nivel = 2
208:                                                Case 3
209:                                                    marca2gen(1)
210:                                                    av1.nivel = 2
211:                                                Case 4
212:                                                    result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
213:                                                    result1.BackColor = Color.Yellow
214:                                                    av1.nivel = 2
215:                                                    End Select
216:                                        Else 'se a qty do 1º aviado for > 150% a do 1º prescrito
217:                                            Hquant(1)
218:                                            av1.nivel = 3
219:                                        End If
220:                                    Else 'se a qty do 1º aviado NÃO contiver apenas algarismos
221:                                        verifQuant(1)
222:                                        av1.nivel = 3
223:                                    End If
224:                                Else  'se forma e dose do 1º aviado NÃO forem todos iguais ao do 1º prescrito
225:                                    If Via(a1array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 1º aviado for diferente da do 1º prescrito
226:                                        apresDif(1)
227:                                    Else
228:                                        If a1array(3).ToString <> p1array(3).ToString Then 'se dose do 1º aviado for diferente da do 1º prescrito
229:                                            DoseDif(1)
230:                                        End If
231:                                    End If
232:                                End If
233:                            End If
234:                        End If
235:                        'Case a2array(1) 'no caso do dci do 1º prescrito ser igual ao do 2º aviado
236:                        If p1array(1) = a2array(1) Then
237:                            If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString And _
                                a2array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 2º aviado forem todos iguais ao do 1º prescrito
239:                                Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
241:                                        OK(2)
242:                                        anular(1)
243:                                        av2.nivel = 0
244:                                    Case 2
245:                                        marca(2)
246:                                        av2.nivel = 2
247:                                    Case 3
248:                                        marca2gen(2)
249:                                        av2.nivel = 2
250:                                    Case 4
251:                                        result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
252:                                        result2.BackColor = Color.Yellow
253:                                        av2.nivel = 2
254:                                    Case Else
255:                                        MsgBox("genlab não devolve comparação")
256:                                        End Select
257:                            Else 'se forma, dose e qty do 2º aviado NÃO forem todos iguais ao do 1º prescrito
258:                                If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString Then
259:                                    'se forma e dose do 2º aviado forem todos iguais ao do 1º prescrito
260:                                    If a2array(4).ToString = p1array(4).ToString Then
261:                                        Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
263:                                                OK(2)
264:                                                anular(1)
265:                                                av2.nivel = 0
266:                                            Case 2
267:                                                marca(2)
268:                                                av2.nivel = 2
269:                                            Case 3
270:                                                marca2gen(2)
271:                                                av2.nivel = 2
272:                                            Case 4
273:                                                result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
274:                                                result2.BackColor = Color.Yellow
275:                                                av2.nivel = 2
276:                                                End Select
277:                                    ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p1array(4).ToString Then 'se a qty do 2º aviado contiver apenas algarismos
278:                                        If a2array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 2º aviado for <= 150% a do 1º prescrito
279:                                            Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
281:                                                    OK(2)
282:                                                    anular(1)
283:                                                    av2.nivel = 2
284:                                                Case 2
285:                                                    marca(2)
286:                                                    av2.nivel = 2
287:                                                Case 3
288:                                                    marca2gen(2)
289:                                                    av2.nivel = 2
290:                                                Case 4
291:                                                    result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
292:                                                    result2.BackColor = Color.Yellow
293:                                                    av2.nivel = 2
294:                                                Case Else
295:                                                    MsgBox("genlab não devolve comparação")
296:                                                    End Select
297:                                        Else 'se a qty do 2º aviado for > 150% a do 1º prescrito
298:                                            Hquant(2)
299:                                            av2.nivel = 3
300:                                        End If
301:                                    Else 'se a qty do 2º aviado NÃO contiver apenas algarismos
302:                                        verifQuant(2)
303:                                        av2.nivel = 3
304:                                    End If
305:                                Else  'se forma e dose do 2º aviado NÃO forem todos iguais ao do 1º prescrito
306:                                    If Via(a2array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 2º aviado for diferente da do 1º prescrito
307:                                        apresDif(2)
308:                                    Else
309:                                        If a2array(3).ToString <> p1array(3).ToString Then 'se dose do 2º aviado for diferente da do 1º prescrito
310:                                            DoseDif(2)
311:                                        End If
312:                                    End If
313:                                End If
314:                            End If
315:                        End If
316:                        If Not p1array(1) = a1array(1) And Not p1array(1) = a2array(1) Then
317:                            'Case Else 'no caso do dci do 1º prescrito NÃO ser igual a nenhum dos dois aviados
318:                            nPresc(1)
319:                            nPresc(2)
320:                        End If
321:                        'End Select
322:                        'End If
323:                        End Select
324:            Else
325:                If IsNothing(a1row) Then
326:                    aviadoNexist(1)
327:                End If
328:                If IsNothing(a2row) Then
329:                    aviadoNexist(2)
330:                End If
331:                msgNexist()
332:            End If
333:            If result1.Text <> "OK" And result2.Text = "OK" Then
334:                If result1.Text <> "F) Embalagem Não Comparticipada" Then
335:                    nPresc(1)
336:                End If
337:            End If
338:            If result2.Text <> "OK" And result1.Text = "OK" Then
339:                If result2.Text <> "F) Embalagem Não Comparticipada" Then
340:                    nPresc(2)
341:                End If
342:            End If
343:        End If
344:        If A = 3 Then
345:            If Not IsNothing(a1row) And Not IsNothing(a2row) And Not IsNothing(a3row) Then 'se o código dos três aviados existe
346:                Dim a1array = a1row.ItemArray
347:                Dim a2array = a2row.ItemArray
348:                Dim a3array = a3row.ItemArray
349:                Select Case p1array(0) 'ver qual o código do 1º prescrito
                        Case a1array(0) 'no caso do código do 1º prescrito ser igual ao do 1º aviado
351:                        If procurado1 = True Then
352:                            OK(1)
353:                            anular(1)
354:                            av1.nivel = 0
355:                        End If
356:                    Case a2array(0) 'no caso do código do 1º prescrito ser igual ao do 2º aviado
357:                        If procurado1 = True Then
358:                            OK(2)
359:                            anular(1)
360:                            av2.nivel = 0
361:                        End If
362:                    Case a3array(0) 'no caso do código do 1º prescrito ser igual ao do 3º aviado
363:                        If procurado1 = True Then
364:                            OK(3)
365:                            anular(1)
366:                            av3.nivel = 0
367:                        End If
368:
369:                    Case Else 'no caso do código do 1º prescrito NÃO ser igual a nenhum dos três aviados
370:                        'If procurado1 = True Then
371:                        'Select Case p1array(1) 'ver qual o dci do 1º prescrito
372:                        If p1array(1) = a1array(1) Then
373:                            'Case a1array(1) 'no caso do dci do 1º prescrito ser igual ao do 1º aviado
374:                            If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString And _
                                a1array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 1º aviado forem todos iguais ao do 1º prescrito
376:                                Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
378:                                        OK(1)
379:                                        anular(1)
380:                                        av1.nivel = 0
381:                                    Case 2
382:                                        marca(1)
383:                                        av1.nivel = 2
384:                                    Case 3
385:                                        marca2gen(1)
386:                                        av1.nivel = 2
387:                                    Case 4
388:                                        result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
389:                                        result1.BackColor = Color.Yellow
390:                                        av1.nivel = 2
391:                                        End Select
392:                            Else 'se forma, dose e qty do 1º aviado NÃO forem todos iguais ao do 1º prescrito
393:                                If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString Then
394:                                    'se forma e dose do 1º aviado forem todos iguais ao do 1º prescrito
395:                                    If a1array(4).ToString = p1array(4).ToString Then
396:                                        Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
398:                                                OK(1)
399:                                                anular(1)
400:                                                av1.nivel = 0
401:                                            Case 2
402:                                                marca(1)
403:                                                av1.nivel = 2
404:                                            Case 3
405:                                                marca2gen(1)
406:                                                av1.nivel = 2
407:                                            Case 4
408:                                                result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
409:                                                result1.BackColor = Color.Yellow
410:                                                av1.nivel = 2
411:                                                End Select
412:                                    ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p1array(4).ToString Then 'se a qty do 1º aviado contiver apenas algarismos
413:                                        If a1array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 1º aviado for <= 150% a do 1º prescrito
414:                                            Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
415:                                                    OK(1)
416:                                                    anular(1)
417:                                                    av1.nivel = 0
418:                                                Case 2
419:                                                    marca(1)
420:                                                    av1.nivel = 2
421:                                                Case 3
422:                                                    marca2gen(1)
423:                                                    av1.nivel = 2
424:                                                Case 4
425:                                                    result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
426:                                                    result1.BackColor = Color.Yellow
427:                                                    av1.nivel = 2
428:                                                    End Select
429:                                        Else 'se a qty do 1º aviado for > 150% a do 1º prescrito
430:                                            Hquant(1)
431:                                            av1.nivel = 3
432:                                        End If
433:                                    Else 'se a qty do 1º aviado NÃO contiver apenas algarismos
434:                                        verifQuant(1)
435:                                        av1.nivel = 3
436:                                    End If
437:                                Else  'se forma e dose do 1º aviado NÃO forem todos iguais ao do 1º prescrito
438:                                    If Via(a1array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 1º aviado for diferente da do 1º prescrito
439:                                        apresDif(1)
440:                                    Else
441:                                        If a1array(3).ToString <> p1array(3).ToString Then 'se dose do 1º aviado for diferente da do 1º prescrito
442:                                            DoseDif(1)
443:                                        End If
444:                                    End If
445:                                End If
446:                            End If
447:                        Else
448:                            dciDif(1)
449:                        End If
450:                        'Case a2array(1) 'no caso do dci do 1º prescrito ser igual ao do 2º aviado
451:                        If p1array(1) = a2array(1) Then
452:                            If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString And _
                                 a2array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 2º aviado forem todos iguais ao do 1º prescrito
453:                                Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(2)
                                            anular(1)
                                            av2.nivel = 0
                                        Case 2
                                            marca(2)
                                            av2.nivel = 2
                                        Case 3
                                            marca2gen(2)
                                            av2.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            av2.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 2º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 2º aviado forem todos iguais ao do 1º prescrito
                                        If a2array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(2)
                                                    anular(1)
                                                    av2.nivel = 0
                                                Case 2
                                                    marca(2)
                                                    av2.nivel = 2
                                                Case 3
                                                    marca2gen(2)
                                                    av2.nivel = 2
                                                Case 4
                                                    result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                                    result2.BackColor = Color.Yellow
                                                    av2.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p1array(4).ToString Then 'se a qty do 2º aviado contiver apenas algarismos
                                            If a2array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 2º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(2)
                                                        anular(1)
                                                        av2.nivel = 0
                                                    Case 2
                                                        marca(2)
                                                        av2.nivel = 2
                                                    Case 3
                                                        marca2gen(2)
                                                        av2.nivel = 2
                                                    Case 4
                                                        result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                                        result2.BackColor = Color.Yellow
                                                        av2.nivel = 2
                                                End Select
                                            Else 'se a qty do 2º aviado for > 150% a do 1º prescrito
                                                Hquant(2)
                                                av2.nivel = 3
                                            End If
                                        Else 'se a qty do 2º aviado NÃO contiver apenas algarismos
                                            verifQuant(2)
                                            av2.nivel = 3
                                        End If
                                    Else  'se forma e dose do 2º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a2array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 2º aviado for diferente da do 1º prescrito
                                            apresDif(2)
                                        Else
                                            If a2array(3).ToString <> p1array(3).ToString Then 'se dose do 2º aviado for diferente da do 1º prescrito
                                                DoseDif(2)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                dciDif(1)
                            End If
                            If p1array(1) = a3array(1) Then
                                'Case a3array(1) 'no caso do dci do 1º prescrito ser igual ao do 3º aviado
                                If Via(a3array(2).ToString) = Via(p1array(2).ToString) And a3array(3).ToString = p1array(3).ToString And _
                                    a3array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 3º aviado forem todos iguais ao do 1º prescrito
                                    Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(3)
                                            anular(1)
                                            av3.nivel = 0
                                        Case 2
                                            marca(3)
                                            av3.nivel = 2
                                        Case 3
                                            marca2gen(3)
                                            av3.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            av3.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 3º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a3array(2).ToString) = Via(p1array(2).ToString) And a3array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 3º aviado forem todos iguais ao do 1º prescrito
                                        If a3array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(3)
                                                    anular(1)
                                                    av3.nivel = 0
                                                Case 2
                                                    marca(3)
                                                    av3.nivel = 2
                                                Case 3
                                                    marca2gen(3)
                                                    av3.nivel = 2
                                                Case 4
                                                    result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                                    result3.BackColor = Color.Yellow
                                                    av3.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p1array(4).ToString Then 'se a qty do 3º aviado contiver apenas algarismos
                                            If a3array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 3º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(3)
                                                        anular(1)
                                                        av3.nivel = 0
                                                    Case 2
                                                        marca(3)
                                                        av3.nivel = 2
                                                    Case 3
                                                        marca2gen(3)
                                                        av3.nivel = 2
                                                    Case 4
                                                        result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                                        result3.BackColor = Color.Yellow
                                                        av3.nivel = 2
                                                End Select
                                            Else 'se a qty do 3º aviado for > 150% a do 1º prescrito
                                                Hquant(3)
                                                av3.nivel = 3
                                            End If
                                        Else 'se a qty do 3º aviado NÃO contiver apenas algarismos
                                            verifQuant(3)
                                            av3.nivel = 3
                                        End If
                                    Else  'se forma e dose do 3º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a3array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 3º aviado for diferente da do 1º prescrito
                                            apresDif(3)
                                        Else
                                            If a3array(3).ToString <> p1array(3).ToString Then 'se dose do 3º aviado for diferente da do 1º prescrito
                                                DoseDif(3)
                                            End If
                                        End If
                                    End If
                                End If
                                If Not p1array(1) = a1array(1) And Not p1array(1) = a2array(1) And Not p1array(1) = a3array(1) Then
                                    'Case Else 'no caso do dci do 1º prescrito NÃO ser igual a nenhum dos três aviados
                                    nPresc(1)
                                    nPresc(2)
                                    nPresc(3)
                                End If
                            Else
                                dciDif(3)
                            End If
                            'End Select
                            'End If
                    End Select
                Else
                    If IsNothing(a1row) Then
                        aviadoNexist(1)
                    End If
                    If IsNothing(a2row) Then
                        aviadoNexist(2)
                    End If
                    If IsNothing(a3row) Then
                        aviadoNexist(3)
                    End If
                    msgNexist()
                End If
                If result1.Text = "OK" Then
                    If result2.Text <> "OK" Then
                        nPresc(2)
                    End If
                    If result3.Text <> "OK" Then
                        nPresc(3)
                    End If
                End If

                If result2.Text = "OK" Then
                    If result1.Text <> "OK" Then
                        nPresc(1)
                    End If
                    If result3.Text <> "OK" Then
                        nPresc(3)
                    End If
                End If

                If result3.Text = "OK" Then
                    If result2.Text <> "OK then" Then
                        nPresc(2)
                    End If
                    If result1.Text <> "OK then" Then
                        nPresc(1)
                    End If
                End If
            End If




            If A = 4 Then
                If Not IsNothing(a1row) And Not IsNothing(a2row) And Not IsNothing(a3row) And Not IsNothing(a4row) Then 'se o código dos quatro aviados existe
                    Dim a1array = a1row.ItemArray
                    Dim a2array = a2row.ItemArray
                    Dim a3array = a3row.ItemArray
                    Dim a4array = a4row.ItemArray
                    Select Case p1array(0) 'ver qual o código do 1º prescrito
                        Case a1array(0) 'no caso do código do 1º prescrito ser igual ao do 1º aviado
                            If procurado1 = True Then
                                OK(1)
                                anular(1)
                                av1.nivel = 0
                            End If
                        Case a2array(0) 'no caso do código do 1º prescrito ser igual ao do 2º aviado
                            If procurado1 = True Then
                                OK(2)
                                anular(1)
                                av2.nivel = 0
                            End If
                        Case a3array(0) 'no caso do código do 1º prescrito ser igual ao do 3º aviado
                            If procurado1 = True Then
                                OK(3)
                                anular(1)
                                av3.nivel = 0
                            End If
                        Case a4array(0) 'no caso do código do 1º prescrito ser igual ao do 4º aviado
                            If procurado1 = True Then
                                OK(4)
                                anular(1)
                                av4.nivel = 0
                            End If
                        Case Else 'no caso do código do 1º prescrito NÃO ser igual a nenhum dos quatro aviados
                            'If procurado1 = True Then
                            'Select Case p1array(1) 'ver qual o dci do 1º prescrito
                            If p1array(1) = a1array(1) Then
                                'Case a1array(1) 'no caso do dci do 1º prescrito ser igual ao do 1º aviado
                                If Via(a1array(2).ToString) = (p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString And _
                                    a1array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 1º aviado forem todos iguais ao do 1º prescrito
                                    Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(1)
                                            anular(1)
                                            av1.nivel = 0
                                        Case 2
                                            marca(1)
                                            av1.nivel = 2
                                        Case 3
                                            marca2gen(1)
                                            av1.nivel = 2
                                        Case 4
                                            result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Yellow
                                            av1.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 1º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a1array(2).ToString) = Via(p1array(2).ToString) And a1array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 1º aviado forem todos iguais ao do 1º prescrito
                                        If a1array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(1)
                                                    anular(1)
                                                    av1.nivel = 0
                                                Case 2
                                                    marca(1)
                                                    av1.nivel = 2
                                                Case 3
                                                    marca2gen(1)
                                                    av1.nivel = 2
                                                Case 4
                                                    result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
                                                    result1.BackColor = Color.Yellow
                                                    av1.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p1array(4).ToString Then 'se a qty do 1º aviado contiver apenas algarismos
                                            If a1array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 1º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a1array(7), p1array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(1)
                                                        anular(1)
                                                        av1.nivel = 0
                                                    Case 2
                                                        marca(1)
                                                        av1.nivel = 2
                                                    Case 3
                                                        marca2gen(1)
                                                        av1.nivel = 2
                                                    Case 4
                                                        result1.Text = "ver se há autorização (de " & p1array(8) & " para " & a1array(8) & ")"
                                                        result1.BackColor = Color.Yellow
                                                        av1.nivel = 2
                                                End Select
                                            Else 'se a qty do 1º aviado for > 150% a do 1º prescrito
                                                Hquant(1)
                                                av1.nivel = 3
                                            End If
                                        Else 'se a qty do 1º aviado NÃO contiver apenas algarismos
                                            verifQuant(1)
                                            av1.nivel = 3
                                        End If
                                    Else  'se forma e dose do 1º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a1array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 1º aviado for diferente da do 1º prescrito
                                            apresDif(1)
                                        Else
                                            If a1array(3).ToString <> p1array(3).ToString Then 'se dose do 1º aviado for diferente da do 1º prescrito
                                                DoseDif(1)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                dciDif(1)
                            End If
                            If p1array(1) = a2array(1) Then
                                'Case a2array(1) 'no caso do dci do 1º prescrito ser igual ao do 2º aviado
                                If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString And _
                                    a2array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 2º aviado forem todos iguais ao do 1º prescrito
                                    Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(2)
                                            anular(1)
                                            av2.nivel = 0
                                        Case 2
                                            marca(2)
                                            av2.nivel = 2
                                        Case 3
                                            marca2gen(2)
                                            av2.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            av2.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 2º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a2array(2).ToString) = Via(p1array(2).ToString) And a2array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 2º aviado forem todos iguais ao do 1º prescrito
                                        If a2array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(2)
                                                    anular(1)
                                                    av2.nivel = 0
                                                Case 2
                                                    marca(2)
                                                    av2.nivel = 2
                                                Case 3
                                                    marca2gen(2)
                                                    av2.nivel = 2
                                                Case 4
                                                    result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                                    result2.BackColor = Color.Yellow
                                                    av2.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p1array(4).ToString Then 'se a qty do 2º aviado contiver apenas algarismos
                                            If a2array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 2º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a2array(7), p1array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(2)
                                                        anular(1)
                                                        av2.nivel = 0
                                                    Case 2
                                                        marca(2)
                                                        av2.nivel = 2
                                                    Case 3
                                                        marca2gen(2)
                                                        av2.nivel = 2
                                                    Case 4
                                                        result2.Text = "ver se há autorização (de " & p1array(8) & " para " & a2array(8) & ")"
                                                        result2.BackColor = Color.Yellow
                                                        av2.nivel = 2
                                                End Select
                                            Else 'se a qty do 2º aviado for > 150% a do 1º prescrito
                                                Hquant(2)
                                                av2.nivel = 3
                                            End If
                                        Else 'se a qty do 2º aviado NÃO contiver apenas algarismos
                                            verifQuant(2)
                                            av2.nivel = 3
                                        End If
                                    Else  'se forma e dose do 2º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a2array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 2º aviado for diferente da do 1º prescrito
                                            apresDif(2)
                                        Else
                                            If a2array(3).ToString <> p1array(3).ToString Then 'se dose do 2º aviado for diferente da do 1º prescrito
                                                DoseDif(2)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                dciDif(2)
                            End If
                            If p1array(1) = a3array(1) Then
                                'Case a3array(1) 'no caso do dci do 1º prescrito ser igual ao do 3º aviado
                                If Via(a3array(2).ToString) = Via(p1array(2).ToString) And a3array(3).ToString = p1array(3).ToString And _
                                    a3array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 3º aviado forem todos iguais ao do 1º prescrito
                                    Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(3)
                                            anular(1)
                                            av3.nivel = 0
                                        Case 2
                                            marca(3)
                                            av3.nivel = 2
                                        Case 3
                                            marca2gen(3)
                                            av3.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            av3.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 3º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a3array(2).ToString) = Via(p1array(2).ToString) And a3array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 3º aviado forem todos iguais ao do 1º prescrito
                                        If a3array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(3)
                                                    anular(1)
                                                    av3.nivel = 0
                                                Case 2
                                                    marca(3)
                                                    av3.nivel = 2
                                                Case 3
                                                    marca2gen(3)
                                                    av3.nivel = 2
                                                Case 4
                                                    result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                                    result3.BackColor = Color.Yellow
                                                    av3.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p1array(4).ToString Then 'se a qty do 3º aviado contiver apenas algarismos
                                            If a3array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 3º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a3array(7), p1array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(3)
                                                        anular(1)
                                                        av3.nivel = 0
                                                    Case 2
                                                        marca(3)
                                                        av3.nivel = 2
                                                    Case 3
                                                        marca2gen(3)
                                                        av3.nivel = 2
                                                    Case 4
                                                        result3.Text = "ver se há autorização (de " & p1array(8) & " para " & a3array(8) & ")"
                                                        result3.BackColor = Color.Yellow
                                                        av3.nivel = 2
                                                End Select
                                            Else 'se a qty do 3º aviado for > 150% a do 1º prescrito
                                                Hquant(3)
                                                av3.nivel = 3
                                            End If
                                        Else 'se a qty do 3º aviado NÃO contiver apenas algarismos
                                            verifQuant(3)
                                            av3.nivel = 3
                                        End If
                                    Else  'se forma e dose do 3º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a3array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 3º aviado for diferente da do 1º prescrito
                                            apresDif(3)
                                        Else
                                            If a3array(3).ToString <> p1array(3).ToString Then 'se dose do 3º aviado for diferente da do 1º prescrito
                                                DoseDif(3)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                dciDif(3)
                            End If
                            If p1array(1) = a4array(1) Then
                                'Case a4array(1) 'no caso do dci do 1º prescrito ser igual ao do 4º aviado
                                If Via(a4array(2).ToString) = Via(p1array(2).ToString) And a4array(3).ToString = p1array(3).ToString And _
                                    a4array(4).ToString = p1array(4).ToString Then 'se forma, dose e qty do 4º aviado forem todos iguais ao do 1º prescrito
                                    Select Case genlab(p1array(7), a4array(7), p1array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            OK(4)
                                            anular(1)
                                            av4.nivel = 0
                                        Case 2
                                            marca(4)
                                            av4.nivel = 2
                                        Case 3
                                            marca2gen(4)
                                            av4.nivel = 2
                                        Case 4
                                            result4.Text = "ver se há autorização (de " & p1array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Yellow
                                            av4.nivel = 2
                                    End Select
                                Else 'se forma, dose e qty do 4º aviado NÃO forem todos iguais ao do 1º prescrito
                                    If Via(a4array(2).ToString) = Via(p1array(2).ToString) And a4array(3).ToString = p1array(3).ToString Then
                                        'se forma e dose do 4º aviado forem todos iguais ao do 1º prescrito
                                        If a4array(4).ToString = p1array(4).ToString Then
                                            Select Case genlab(p1array(7), a4array(7), p1array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                                Case 1
                                                    OK(4)
                                                    anular(1)
                                                    av4.nivel = 0
                                                Case 2
                                                    marca(4)
                                                    av4.nivel = 2
                                                Case 3
                                                    marca2gen(4)
                                                    av4.nivel = 2
                                                Case 4
                                                    result4.Text = "ver se há autorização (de " & p1array(8) & " para " & a4array(8) & ")"
                                                    result4.BackColor = Color.Yellow
                                                    av4.nivel = 2
                                            End Select
                                        ElseIf IsNumeric(a4array(4).ToString) And Not a4array(4).ToString = p1array(4).ToString Then 'se a qty do 4º aviado contiver apenas algarismos
                                            If a4array(4).ToString <= 1.5 * p1array(4).ToString Then 'se a qty do 4º aviado for <= 150% a do 1º prescrito
                                                Select Case genlab(p1array(7), a4array(7), p1array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                                    Case 1
                                                        OK(4)
                                                        anular(1)
                                                        av4.nivel = 0
                                                    Case 2
                                                        marca(4)
                                                        av4.nivel = 2
                                                    Case 3
                                                        marca2gen(4)
                                                        av4.nivel = 2
                                                    Case 4
                                                        result4.Text = "ver se há autorização (de " & p1array(8) & " para " & a4array(8) & ")"
                                                        result4.BackColor = Color.Yellow
                                                        av4.nivel = 2
                                                End Select
                                            Else 'se a qty do 4º aviado for > 150% a do 1º prescrito
                                                Hquant(4)
                                                av4.nivel = 3
                                            End If
                                        Else 'se a qty do 4º aviado NÃO contiver apenas algarismos
                                            verifQuant(4)
                                            av4.nivel = 3
                                        End If
                                    Else  'se forma e dose do 4º aviado NÃO forem todos iguais ao do 1º prescrito
                                        If Via(a4array(2).ToString) <> Via(p1array(2).ToString) Then 'se forma do 4º aviado for diferente da do 1º prescrito
                                            apresDif(4)
                                        Else
                                            If a4array(3).ToString <> p1array(3).ToString Then 'se dose do 4º aviado for diferente da do 1º prescrito
                                                DoseDif(4)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                dciDif(4)
                            End If
                            If Not p1array(1) = a1array(1) And Not p1array(1) = a2array(1) And Not p1array(1) = a3array(1) And Not p1array(1) = a4array(1) Then
                                'Case Else 'no caso do dci do 1º prescrito NÃO ser igual a nenhum dos três aviados
                                nPresc(1)
                                nPresc(2)
                                nPresc(3)
                                nPresc(4)
                                'End Select
                            End If
                    End Select
                Else
                    If IsNothing(a1row) Then
                        aviadoNexist(1)
                    End If
                    If IsNothing(a2row) Then
                        aviadoNexist(2)
                    End If
                    If IsNothing(a3row) Then
                        aviadoNexist(3)
                    End If
                    If IsNothing(a4row) Then
                        aviadoNexist(4)
                    End If
                    msgNexist()
                End If
                If result1.Text = "OK" Then
                    If result2.Text <> "OK" Then
                        nPresc(2)
                    End If
                    If result3.Text <> "OK" Then
                        nPresc(3)
                    End If
                    If result4.Text <> "OK" Then
                        nPresc(4)
                    End If
                End If

                If result2.Text = "OK" Then
                    If result1.Text <> "OK" Then
                        nPresc(1)
                    End If
                    If result3.Text <> "OK" Then
                        nPresc(3)
                    End If
                    If result4.Text <> "OK" Then
                        nPresc(4)
                    End If
                End If

                If result3.Text = "OK" Then
                    If result2.Text <> "OK then" Then
                        nPresc(2)
                    End If
                    If result1.Text <> "OK then" Then
                        nPresc(1)
                    End If
                    If result4.Text <> "OK then" Then
                        nPresc(4)
                    End If
                End If
                If result4.Text = "OK" Then
                    If result1.Text <> "OK then" Then
                        nPresc(1)
                    End If
                    If result2.Text <> "OK then" Then
                        nPresc(2)
                    End If
                    If result3.Text <> "OK then" Then
                        nPresc(3)
                    End If
                End If
            End If
        Else
            prescNexist(1)
        End If
        If result1.Text = "" Or result1.Text = "G) DCI não prescrito" Then
            nPresc(1)
        End If
        If aviam2.Text <> 0 And (result2.Text = "" Or result2.Text = "G) DCI não prescrito") Then
            nPresc(2)
        End If
        If aviam3.Text <> 0 And (result3.Text = "" Or result3.Text = "G) DCI não prescrito") Then
            nPresc(3)
        End If
        If aviam4.Text <> 0 And (result4.Text = "" Or result4.Text = "G) DCI não prescrito") Then
            nPresc(4)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Sub p1a1()
        On Error GoTo MOSTRARERRO
1:      If Not IsNothing(a1row) And Not IsNothing(p1row) Then 'se o código do 1º aviado e 1º prescrito existem
2:          Dim a1array = a1row.ItemArray
3:          Dim p1Array = p1row.ItemArray
4:          If a1array(0).ToString = p1Array(0).ToString Then 'se o código do 1º aviado for igual ao do 1º prescrito
5:              a1p1.nivel = 0
6:              a1p1.resultado = 0
7:              descodificar(0, 1)
8:          Else
9:              If a1array(1).ToString = p1Array(1).ToString Then 'se o dci do 1º aviado for igual ao do 1º prescrito
10:                 If Via(a1array(2).ToString) = Via(p1Array(2).ToString) Then 'se a apresentação do 1º aviado for igual ao do 1º prescrito
11:                     If a1array(3).ToString = p1Array(3).ToString Then 'se a dosagem do 1º aviado for igual ao do 1º prescrito
12:                         If a1array(4).ToString = p1Array(4).ToString Then
13:                             a1p2.mostrado = True
14:                             a1p1.mostrado = True
15:                             a1p3.mostrado = True
16:                             a1p4.mostrado = True
17:                             Select Case genlab(p1Array(7), a1array(7), p1Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
18:                                     descodificar(0, 1)
19:                                     anular(1)
20:                                     a1p1.nivel = 0
21:                                 Case 2
22:                                     descodificar(8, 1)
23:                                     a1p1.nivel = 2
24:                                 Case 3
25:                                     descodificar(7, 1)
26:                                     a1p1.nivel = 2
27:                                 Case 4
28:                                     result1.Text = "ver se há autorização (de " & p1Array(8) & " para " & a1array(8) & ")"
29:                                     result1.BackColor = Color.Yellow
30:                                     a1p1.nivel = 2
31:                                 Case Else
32:                                     MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
33:                                     End Select
34:                         ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p1Array(4).ToString Then 'se a quantidade do 1º aviadocontiver só algarismos
35:                             a1p2.mostrado = True
36:                             a1p1.mostrado = True
37:                             a1p3.mostrado = True
38:                             a1p4.mostrado = True
39:                             If a1array(4).ToString <= 1.5 * p1Array(4).ToString Then 'se a quantidade do 1º aviado for inferior a 150% do 1º prescrito
40:                                 a1p1.nivel = 2
41:                                 Select Case genlab(p1Array(7), a1array(7), p1Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
43:                                         descodificar(0, 1)
44:                                         anular(1)
45:                                         a1p1.nivel = 0
46:                                     Case 2
47:                                         descodificar(8, 1)
48:                                         a1p1.nivel = 2
49:                                     Case 3
50:                                         descodificar(7, 1)
51:                                         a1p1.nivel = 2
52:                                     Case 4
53:                                         result1.Text = "ver se há autorização (de " & p1Array(8) & " para " & a1array(8) & ")"
54:                                         result1.BackColor = Color.Yellow
55:                                         a1p1.nivel = 2
56:                                     Case Else
57:                                         MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
58:                                         End Select
59:                             Else 'se a quantidade do 1º aviado NÃO for inferior a 150% do 1º prescrito
60:                                 Select Case genlab(p1Array(7), a1array(7), p1Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
62:                                         descodificar(2, 1)
63:                                         anular(1)
64:                                         a1p1.nivel = 2 'nivel do h normal
65:                                     Case 2
66:                                         result1.Text = "H) + Aviamento de marca (" & a1array(0) & ")"
67:                                         result1.BackColor = Color.Red
68:                                         a1p1.nivel = 2 'nivel daqui
69:                                         MsgBox("troca no " & a1array(0))
70:                                     Case 3
71:                                         result1.Text = "H) + marca (" & p1Array(0) & ") trocado para genérico (" & a1array(0) & ")"
72:                                         result1.BackColor = Color.Red
73:                                         a1p1.nivel = 2 'nivel daqui
74:                                         MsgBox("troca no " & a1array(0))
75:                                     Case 4
76:                                         result1.Text = "H) + ver se há autorização (de " & p1Array(8) & " para " & a1array(8) & ")"
77:                                         result1.BackColor = Color.Red
78:                                         a1p1.nivel = 2 'nivel daqui
79:                                         MsgBox("troca no " & a1array(0))
80:                                     Case Else
81:                                         MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
82:                                         a1p1.nivel = 3
83:                                         End Select
84:                                 a1p2.mostrado = False
85:                                 a1p1.mostrado = False
86:                                 a1p3.mostrado = False
87:                                 a1p4.mostrado = False
88:                                 a1p1.resultado = 2
89:                                 a1p1.nivel = 2.5
90:                             End If
91:                         Else 'se a quantidade do 1º aviado NÃO contiver só algarismos
92:                             a1p1.resultado = 1
93:                             a1p1.nivel = 3
94:                             descodificar(1, 1)
95:                         End If
96:                         av1.nivel = a1p1.nivel
97:                     Else 'se a dosagem do 1º aviado NÃO for igual ao do 1º prescrito
98:                         a1p1.resultado = 3
99:                         a1p1.nivel = 4
100:                        descodificar(3, 1)
101:                    End If
102:                Else 'se a apresentação do 1º aviado NÃO for igual ao do 1º prescrito
103:                    a1p1.resultado = 4
104:                    a1p1.nivel = 5
105:                    'descodificar(4, 1)
106:                End If
107:            Else 'se o dci do 1º aviado NÃO for igual ao do 1º prescrito
108:                a1p1.resultado = 5
109:                a1p1.nivel = 6
110:                'descodificar(5, 1)
111:            End If
112:        End If
113:
114:    Else
115:        If IsNothing(p1row) Then
116:            prescNexist(1)
117:        End If
118:        If IsNothing(a1row) Then
119:            aviadoNexist(1)
120:            msgNexist()
121:        End If
122:    End If
123:    If result1.Text = "" Then
124:        nPresc(1)
125:    End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p1a2()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) And Not IsNothing(p1row) Then 'se o código do 2º aviado e 1º prescrito existem
            Dim a2array = a2row.ItemArray
            Dim p1Array = p1row.ItemArray
            If a2array(0).ToString = p1Array(0).ToString Then 'se o código do 2º aviado for igual ao do 1º prescrito
                a2p1.nivel = 0
                a2p1.resultado = 0
                descodificar(0, 2)
            Else
                If a2array(1).ToString = p1Array(1).ToString Then 'se o dci do 2º aviado for igual ao do 1º prescrito
                    If Via(a2array(2).ToString) = Via(p1Array(2).ToString) Then 'se a apresentação do 2º aviado for igual ao do 1º prescrito
                        If a2array(3).ToString = p1Array(3).ToString Then 'se a dosagem do 2º aviado for igual ao do 1º prescrito
                            If a2array(4).ToString = p1Array(4).ToString Then
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                Select Case genlab(p1Array(7), a2array(7), p1Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 2)
                                        anular(1)
                                        a2p1.nivel = 0
                                    Case 2
                                        descodificar(8, 2)
                                        a1p1.nivel = 2
                                    Case 3
                                        descodificar(7, 2)
                                        a2p1.nivel = 2
                                    Case 4
                                        result2.Text = "ver se há autorização (de " & p1Array(8) & " para " & a2array(8) & ")"
                                        result2.BackColor = Color.Yellow
                                        a2p1.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p1Array(4).ToString Then 'se a quantidade do 2º aviado contiver só algarismos
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                If a2array(4).ToString <= 1.5 * p1Array(4).ToString Then 'se a quantidade do 2º aviado for inferior a 150% do 1º prescrito
                                    a2p1.nivel = 2
                                    Select Case genlab(p1Array(7), a2array(7), p1Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 2)
                                            anular(1)
                                            a2p1.nivel = 0
                                        Case 2
                                            descodificar(8, 2)
                                            a2p1.nivel = 2
                                        Case 3
                                            descodificar(7, 2)
                                            a2p1.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p1Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            a2p1.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 2º aviado NÃO for inferior a 150% do 1º prescrito
                                    Select Case genlab(p1Array(7), a2array(7), p1Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 2)
                                            anular(1)
                                            a2p1.nivel = 2 'nivel do h normal
                                        Case 2
                                            result2.Text = "H) + Aviamento de marca (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 3
                                            result2.Text = "H) + marca (" & p1Array(0) & ") trocado para genérico (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 4
                                            result2.Text = "H) + ver se há autorização (de " & p1Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Red
                                            a2p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a2p1.nivel = 3
                                    End Select
                                    a2p2.mostrado = False
                                    a2p1.mostrado = False
                                    a2p3.mostrado = False
                                    a2p4.mostrado = False
                                    a2p1.resultado = 2
                                    a2p1.nivel = 2.5
                                End If
                            Else 'se a quantidade do 2º aviado NÃO contiver só algarismos
                                a2p1.resultado = 1
                                a2p1.nivel = 3
                                descodificar(1, 2)
                            End If
                            av2.nivel = a2p1.nivel
                        Else 'se a dosagem do 2º aviado NÃO for igual ao do 1º prescrito
                            a2p1.resultado = 3
                            a2p1.nivel = 4
                            descodificar(3, 2)
                        End If
                    Else 'se a apresentação do 2º aviado NÃO for igual ao do 1º prescrito
                        a2p1.resultado = 4
                        a2p1.nivel = 5
                        'descodificar(4, 2)
                    End If
                Else 'se o dci do 2º aviado NÃO for igual ao do 1º prescrito
                    a2p1.resultado = 5
                    a2p1.nivel = 6
                    'descodificar(5, 2)
                End If
            End If

        Else
            If IsNothing(p1row) Then
                prescNexist(1)
            End If
            If IsNothing(a2row) Then
                aviadoNexist(2)
                msgNexist()
            End If
        End If
        If result2.Text = "" Then
            nPresc(2)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p1a3()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) And Not IsNothing(p1row) Then 'se o código do 3º aviado e 1º prescrito existem
            Dim a3array = a3row.ItemArray
            Dim p1Array = p1row.ItemArray
            If a3array(0).ToString = p1Array(0).ToString Then 'se o código do 3º aviado for igual ao do 1º prescrito
                a3p1.nivel = 0
                a3p1.resultado = 0
                descodificar(0, 3)
            Else
                If a3array(1).ToString = p1Array(1).ToString Then 'se o dci do 3º aviado for igual ao do 1º prescrito
                    If Via(a3array(2).ToString) = Via(p1Array(2).ToString) Then 'se a apresentação do 3º aviado for igual ao do 1º prescrito
                        If a3array(3).ToString = p1Array(3).ToString Then 'se a dosagem do 3º aviado for igual ao do 1º prescrito
                            If a3array(4).ToString = p1Array(4).ToString Then
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                Select Case genlab(p1Array(7), a3array(7), p1Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 3)
                                        anular(1)
                                        a3p1.nivel = 0
                                    Case 2
                                        descodificar(8, 3)
                                        a3p1.nivel = 2
                                    Case 3
                                        descodificar(7, 3)
                                        a3p1.nivel = 2
                                    Case 4
                                        result3.Text = "ver se há autorização (de " & p1Array(8) & " para " & a3array(8) & ")"
                                        result3.BackColor = Color.Yellow
                                        a3p1.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p1Array(4).ToString Then 'se a quantidade do 3º aviado contiver só algarismos
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                If a3array(4).ToString <= 1.5 * p1Array(4).ToString Then 'se a quantidade do 3º aviado for inferior a 150% do 1º prescrito
                                    a3p1.nivel = 2
                                    Select Case genlab(p1Array(7), a3array(7), p1Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 3)
                                            anular(1)
                                            a3p1.nivel = 0
                                        Case 2
                                            descodificar(8, 3)
                                            a3p1.nivel = 2
                                        Case 3
                                            descodificar(7, 3)
                                            a3p1.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p1Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            a3p1.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 3º aviado NÃO for inferior a 150% do 1º prescrito
                                    Select Case genlab(p1Array(7), a3array(7), p1Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 3)
                                            anular(1)
                                            a3p1.nivel = 2 'nivel do h normal
                                        Case 2
                                            result3.Text = "H) + Aviamento de marca (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 3
                                            result3.Text = "H) + marca (" & p1Array(0) & ") trocado para genérico (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 4
                                            result3.Text = "H) + ver se há autorização (de " & p1Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Red
                                            a3p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a3p1.nivel = 3

                                    End Select
                                    a3p2.mostrado = False
                                    a3p3.mostrado = False
                                    a3p4.mostrado = False
                                    a3p1.mostrado = False
                                    a3p1.resultado = 2
                                    a3p1.nivel = 2.5
                                End If
                            Else 'se a quantidade do 3º aviado NÃO contiver só algarismos
                                a3p1.resultado = 1
                                a3p1.nivel = 3
                                descodificar(1, 3)
                            End If
                            av3.nivel = a3p1.nivel
                        Else 'se a dosagem do 3º aviado NÃO for igual ao do 1º prescrito
                            a3p1.resultado = 3
                            a3p1.nivel = 4
                            descodificar(3, 3)
                        End If
                    Else 'se a apresentação do 3º aviado NÃO for igual ao do 1º prescrito
                        a3p1.resultado = 4
                        a3p1.nivel = 5
                        'descodificar(4, 3)
                    End If
                Else 'se o dci do 3º aviado NÃO for igual ao do 1º prescrito
                    a3p1.resultado = 5
                    a3p1.nivel = 6
                    'descodificar(5, 3)
                End If
            End If

        Else
            If IsNothing(p1row) Then
                prescNexist(1)
            End If
            If IsNothing(a3row) Then
                aviadoNexist(3)
                msgNexist()
            End If
        End If
        If result3.Text = "" Then
            nPresc(3)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p1a4()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) And Not IsNothing(p1row) Then 'se o código do 4º aviado e 1º prescrito existem
            Dim a4array = a4row.ItemArray
            Dim p1Array = p1row.ItemArray
            If a4array(0).ToString = p1Array(0).ToString Then 'se o código do 4º aviado for igual ao do 1º prescrito
                a4p1.nivel = 0
                a4p1.resultado = 0
                descodificar(0, 4)
            Else
                If a4array(1).ToString = p1Array(1).ToString Then 'se o dci do 4º aviado for igual ao do 1º prescrito
                    If Via(a4array(2).ToString) = Via(p1Array(2).ToString) Then 'se a apresentação do 4º aviado for igual ao do 1º prescrito
                        If a4array(3).ToString = p1Array(3).ToString Then 'se a dosagem do 4º aviado for igual ao do 1º prescrito
                            If a4array(4).ToString = p1Array(4).ToString Then
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                Select Case genlab(p1Array(7), a4array(7), p1Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 4)
                                        anular(1)
                                        a4p1.nivel = 0
                                    Case 2
                                        descodificar(8, 4)
                                        a4p1.nivel = 2
                                    Case 3
                                        descodificar(7, 4)
                                        a4p1.nivel = 2
                                    Case 4
                                        result4.Text = "ver se há autorização (de " & p1Array(8) & " para " & a4array(8) & ")"
                                        result4.BackColor = Color.Yellow
                                        a4p1.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a4array(4).ToString) And Not a4array(4).ToString = p1Array(4).ToString Then 'se a quantidade do 4º aviado contiver só algarismos
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                If a4array(4).ToString <= 1.5 * p1Array(4).ToString Then 'se a quantidade do 4º aviado for inferior a 150% do 1º prescrito
                                    a4p1.nivel = 2
                                    Select Case genlab(p1Array(7), a4array(7), p1Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 4)
                                            anular(1)
                                            a4p1.nivel = 0
                                        Case 2
                                            descodificar(8, 4)
                                            a4p1.nivel = 2
                                        Case 3
                                            descodificar(7, 4)
                                            a4p1.nivel = 2
                                        Case 4
                                            result4.Text = "ver se há autorização (de " & p1Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Yellow
                                            a4p1.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select

                                Else 'se a quantidade do 4º aviado NÃO for inferior a 150% do 1º prescrito
                                    Select Case genlab(p1Array(7), a4array(7), p1Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 4)
                                            anular(1)
                                            a4p1.nivel = 2 'nivel do h normal
                                        Case 2
                                            result4.Text = "H) + Aviamento de marca (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 3
                                            result4.Text = "H) + marca (" & p1Array(0) & ") trocado para genérico (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 4
                                            result4.Text = "H) + ver se há autorização (de " & p1Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Red
                                            a4p1.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a4p1.nivel = 3
                                    End Select
                                    a4p2.mostrado = False
                                    a4p1.mostrado = False
                                    a4p3.mostrado = False
                                    a4p4.mostrado = False
                                    a4p1.resultado = 2
                                    a4p1.nivel = 2.5
                                End If

                            Else 'se a quantidade do 4º aviado NÃO contiver só algarismos
                                a4p1.resultado = 1
                                a4p1.nivel = 3
                                descodificar(1, 4)
                            End If
                            av4.nivel = a4p1.nivel
                        Else 'se a dosagem do 4º aviado NÃO for igual ao do 1º prescrito
                            a4p1.resultado = 3
                            a4p1.nivel = 4
                            descodificar(3, 4)
                        End If
                    Else 'se a apresentação do 4º aviado NÃO for igual ao do 1º prescrito
                        a4p1.resultado = 4
                        a4p1.nivel = 5
                        'descodificar(4, 4)
                    End If

                Else 'se o dci do 4º aviado NÃO for igual ao do 1º prescrito
                    a4p1.resultado = 5
                    a4p1.nivel = 6
                    'descodificar(5, 4)
                End If
            End If
        Else
            If IsNothing(p1row) Then
                prescNexist(1)
            End If
            If IsNothing(a4row) Then
                aviadoNexist(4)
                msgNexist()
            End If
        End If
        If result4.Text = "" Then
            nPresc(4)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p2a1()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) And Not IsNothing(p2row) Then 'se o código do 1º aviado e 2º prescrito existem
            Dim a1array = a1row.ItemArray
            Dim p2Array = p2row.ItemArray
            If a1array(0).ToString = p2Array(0).ToString Then 'se o código do 1º aviado for igual ao do 2º prescrito
                a1p2.nivel = 0
                a1p2.resultado = 0
                descodificar(0, 1)
            Else
                If a1array(1).ToString = p2Array(1).ToString Then 'se o dci do 1º aviado for igual ao do 2º prescrito
                    If Via(a1array(2).ToString) = Via(p2Array(2).ToString) Then 'se a apresentação do 1º aviado for igual ao do 2º prescrito
                        If a1array(3).ToString = p2Array(3).ToString Then 'se a dosagem do 1º aviado for igual ao do 2º prescrito
                            If a1array(4).ToString = p2Array(4).ToString Then
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                Select Case genlab(p2Array(7), a1array(7), p2Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 1)
                                        anular(2)
                                        a1p2.nivel = 0
                                    Case 2
                                        descodificar(8, 1)
                                        a1p2.nivel = 2
                                    Case 3
                                        descodificar(7, 1)
                                        a1p2.nivel = 2
                                    Case 4
                                        result1.Text = "ver se há autorização (de " & p2Array(8) & " para " & a1array(8) & ")"
                                        result1.BackColor = Color.Yellow
                                        a1p2.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p2Array(4).ToString Then 'se a quantidade do 1º aviado NÃO contiver só algarismos
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                If a1array(4).ToString <= 1.5 * p2Array(4).ToString Then 'se a quantidade do 1º aviado for inferior a 150% do 2º prescrito
                                    a1p2.nivel = 2
                                    Select Case genlab(p2Array(7), a1array(7), p2Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 1)
                                            anular(2)
                                            a1p2.nivel = 0
                                        Case 2
                                            descodificar(8, 1)
                                            a1p2.nivel = 2
                                        Case 3
                                            descodificar(7, 1)
                                            a1p2.nivel = 2
                                        Case 4
                                            result1.Text = "ver se há autorização (de " & p2Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Yellow
                                            a1p2.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 1º aviado NÃO for inferior a 150% do 2º prescrito
                                    Select Case genlab(p2Array(7), a1array(7), p2Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 1)
                                            anular(2)
                                            a1p2.nivel = 2 'nivel do h normal
                                        Case 2
                                            result1.Text = "H) + Aviamento de marca (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 3
                                            result1.Text = "H) + marca (" & p2Array(0) & ") trocado para genérico (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 4
                                            result1.Text = "H) + ver se há autorização (de " & p2Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Red
                                            a1p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a1p2.nivel = 3
                                    End Select
                                    a1p2.mostrado = False
                                    a1p1.mostrado = False
                                    a1p3.mostrado = False
                                    a1p4.mostrado = False
                                    a1p2.resultado = 2
                                    a1p2.nivel = 2.5
                                End If
                            Else 'se a quantidade do 1º aviado NÃO contiver só algarismos
                                a1p2.resultado = 1
                                a1p2.nivel = 3
                                descodificar(1, 1)
                            End If
                            av1.nivel = a1p2.nivel
                        Else 'se a dosagem do 1º aviado NÃO for igual ao do 2º prescrito
                            a1p2.resultado = 3
                            a1p2.nivel = 4
                            descodificar(3, 1)
                        End If
                    Else 'se a apresentação do 1º aviado NÃO for igual ao do 2º prescrito
                        a1p2.resultado = 4
                        a1p2.nivel = 5
                        'descodificar(4, 1)
                    End If
                Else 'se o dci do 1º aviado NÃO for igual ao do 2º prescrito
                    a1p2.resultado = 5
                    a1p2.nivel = 6
                    'descodificar(5, 1)
                End If
            End If

        Else
            If IsNothing(p2row) Then
                prescNexist(2)
            End If
            If IsNothing(a1row) Then
                aviadoNexist(1)
                msgNexist()
            End If
        End If
        If result1.Text = "" Then
            nPresc(1)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p2a2()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) And Not IsNothing(p2row) Then 'se o código do 2º aviado e 2º prescrito existem
            Dim a2array = a2row.ItemArray
            Dim p2Array = p2row.ItemArray
            If a2array(0).ToString = p2Array(0).ToString Then 'se o código do 2º aviado for igual ao do 2º prescrito
                a2p2.nivel = 0
                a2p2.resultado = 0
                descodificar(0, 2)
            Else
                If a2array(1).ToString = p2Array(1).ToString Then 'se o dci do 2º aviado for igual ao do 2º prescrito
                    If Via(a2array(2).ToString) = Via(p2Array(2).ToString) Then 'se a apresentação do 2º aviado for igual ao do 2º prescrito
                        If a2array(3).ToString = p2Array(3).ToString Then 'se a dosagem do 2º aviado for igual ao do 2º prescrito
                            If a2array(4).ToString = p2Array(4).ToString Then
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                Select Case genlab(p2Array(7), a2array(7), p2Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 2)
                                        anular(2)
                                        a2p2.nivel = 0
                                    Case 2
                                        descodificar(8, 2)
                                        a2p2.nivel = 2
                                    Case 3
                                        descodificar(7, 2)
                                        a2p2.nivel = 2
                                    Case 4
                                        result2.Text = "ver se há autorização (de " & p2Array(8) & " para " & a2array(8) & ")"
                                        result2.BackColor = Color.Yellow
                                        a2p2.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p2Array(4).ToString Then 'se a quantidade do 2º aviado contiver só algarismos
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                If a2array(4).ToString <= 1.5 * p2Array(4).ToString Then 'se a quantidade do 2º aviado for inferior a 150% do 2º prescrito
                                    a2p2.nivel = 2
                                    Select Case genlab(p2Array(7), a2array(7), p2Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 2)
                                            anular(2)
                                            a2p2.nivel = 0
                                        Case 2
                                            descodificar(8, 2)
                                            a2p2.nivel = 2
                                        Case 3
                                            descodificar(7, 2)
                                            a2p2.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p2Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            a2p2.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 2º aviado NÃO for inferior a 150% do 2º prescrito
                                    Select Case genlab(p2Array(7), a2array(7), p2Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 2)
                                            anular(2)
                                            a2p2.nivel = 2 'nivel do h normal
                                        Case 2
                                            result2.Text = "H) + Aviamento de marca (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 3
                                            result2.Text = "H) + marca (" & p2Array(0) & ") trocado para genérico (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 4
                                            result2.Text = "H) + ver se há autorização (de " & p2Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Red
                                            a2p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a2p2.nivel = 3
                                    End Select
                                    a2p2.mostrado = False
                                    a2p1.mostrado = False
                                    a2p3.mostrado = False
                                    a2p4.mostrado = False
                                    a2p2.resultado = 2
                                    a2p2.nivel = 2.5
                                End If
                            Else 'se a quantidade do 2º aviado NÃO contiver só algarismos
                                a2p2.resultado = 1
                                a2p2.nivel = 3
                                descodificar(1, 2)
                            End If
                            av2.nivel = a2p2.nivel
                        Else 'se a dosagem do 2º aviado NÃO for igual ao do 2º prescrito
                            a2p2.resultado = 3
                            a2p2.nivel = 4
                            descodificar(3, 2)
                        End If
                    Else 'se a apresentação do 2º aviado NÃO for igual ao do 2º prescrito
                        a2p2.resultado = 4
                        a2p2.nivel = 5
                        'descodificar(4, 2)
                    End If
                Else 'se o dci do 2º aviado NÃO for igual ao do 2º prescrito
                    a2p2.resultado = 5
                    a2p2.nivel = 6
                    'descodificar(5, 2)
                End If
            End If

        Else
            If IsNothing(p2row) Then
                prescNexist(2)
            End If
            If IsNothing(a2row) Then
                aviadoNexist(2)
                msgNexist()
            End If
        End If
        If result2.Text = "" Then
            nPresc(2)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p2a3()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) And Not IsNothing(p2row) Then 'se o código do 3º aviado e 2º prescrito existem
            Dim a3array = a3row.ItemArray
            Dim p2Array = p2row.ItemArray
            If a3array(0).ToString = p2Array(0).ToString Then 'se o código do 3º aviado for igual ao do 2º prescrito
                a3p2.nivel = 0
                a3p2.resultado = 0
                descodificar(0, 3)
            Else
                If a3array(1).ToString = p2Array(1).ToString Then 'se o dci do 3º aviado for igual ao do 2º prescrito
                    If Via(a3array(2).ToString) = Via(p2Array(2).ToString) Then 'se a apresentação do 3º aviado for igual ao do 2º prescrito
                        If a3array(3).ToString = p2Array(3).ToString Then 'se a dosagem do 3º aviado for igual ao do 2º prescrito
                            If a3array(4).ToString = p2Array(4).ToString Then
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                Select Case genlab(p2Array(7), a3array(7), p2Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 3)
                                        anular(2)
                                        a3p2.nivel = 0
                                    Case 2
                                        descodificar(8, 3)
                                        a3p2.nivel = 2
                                    Case 3
                                        descodificar(7, 3)
                                        a3p2.nivel = 2
                                    Case 4
                                        result3.Text = "ver se há autorização (de " & p2Array(8) & " para " & a3array(8) & ")"
                                        result3.BackColor = Color.Yellow
                                        a3p2.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p2Array(4).ToString Then 'se a quantidade do 3º aviado contiver só algarismos
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                If a3array(4).ToString <= 1.5 * p2Array(4).ToString Then 'se a quantidade do 3º aviado for inferior a 150% do 2º prescrito
                                    a3p2.nivel = 2
                                    Select Case genlab(p2Array(7), a3array(7), p2Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 3)
                                            anular(2)
                                            a3p2.nivel = 0
                                        Case 2
                                            descodificar(8, 3)
                                            a3p2.nivel = 2
                                        Case 3
                                            descodificar(7, 3)
                                            a3p2.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p2Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            a3p2.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 3º aviado NÃO for inferior a 150% do 2º prescrito
                                    Select Case genlab(p2Array(7), a3array(7), p2Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 3)
                                            anular(2)
                                            a3p2.nivel = 2 'nivel do h normal
                                        Case 2
                                            result3.Text = "H) + Aviamento de marca (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 3
                                            result3.Text = "H) + marca (" & p2Array(0) & ") trocado para genérico (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 4
                                            result3.Text = "H) + ver se há autorização (de " & p2Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Red
                                            a3p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a3p2.nivel = 3
                                    End Select
                                    a3p2.mostrado = False
                                    a3p1.mostrado = False
                                    a3p3.mostrado = False
                                    a3p4.mostrado = False
                                    a3p2.resultado = 2
                                    a3p2.nivel = 2.5
                                End If
                            Else 'se a quantidade do 3º aviado NÃO contiver só algarismos
                                a3p2.resultado = 1
                                a3p2.nivel = 3
                                descodificar(1, 3)
                            End If
                            av3.nivel = a3p2.nivel
                        Else 'se a dosagem do 3º aviado NÃO for igual ao do 2º prescrito
                            a3p2.resultado = 3
                            a3p2.nivel = 4
                            descodificar(3, 3)
                        End If
                    Else 'se a apresentação do 3º aviado NÃO for igual ao do 2º prescrito
                        a3p2.resultado = 4
                        a3p2.nivel = 5
                        'descodificar(4, 3)
                    End If
                Else 'se o dci do 3º aviado NÃO for igual ao do 2º prescrito
                    a3p2.resultado = 5
                    a3p2.nivel = 6
                    'descodificar(5, 3)
                End If
            End If
        Else
            If IsNothing(p2row) Then
                prescNexist(2)
            End If
            If IsNothing(a3row) Then
                aviadoNexist(3)
                msgNexist()
            End If
        End If
        If result3.Text = "" Then
            nPresc(3)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p2a4()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) And Not IsNothing(p2row) Then 'se o código do 4º aviado e 2º prescrito existem
            Dim a4array = a4row.ItemArray
            Dim p2Array = p2row.ItemArray
            If a4array(0).ToString = p2Array(0).ToString Then 'se o código do 4º aviado for igual ao do 2º prescrito
                a4p2.nivel = 0
                a4p2.resultado = 0
                descodificar(0, 4)
            Else
                If a4array(1).ToString = p2Array(1).ToString Then 'se o dci do 4º aviado for igual ao do 2º prescrito
                    If Via(a4array(2).ToString) = Via(p2Array(2).ToString) Then 'se a apresentação do 4º aviado for igual ao do 2º prescrito
                        If a4array(3).ToString = p2Array(3).ToString Then 'se a dosagem do 4º aviado for igual ao do 2º prescrito
                            If a4array(4).ToString = p2Array(4).ToString Then
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                Select Case genlab(p2Array(7), a4array(7), p2Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 4)
                                        anular(2)
                                        a4p2.nivel = 0
                                    Case 2
                                        descodificar(8, 4)
                                        a4p2.nivel = 2
                                    Case 3
                                        descodificar(7, 4)
                                        a4p2.nivel = 2
                                    Case 4
                                        result4.Text = "ver se há autorização (de " & p2Array(8) & " para " & a4array(8) & ")"
                                        result4.BackColor = Color.Yellow
                                        a4p2.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a4array(4).ToString) And Not a4array(4).ToString = p2Array(4).ToString Then 'se a quantidade do 4º aviado contiver só algarismos
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                If a4array(4).ToString <= 1.5 * p2Array(4).ToString Then 'se a quantidade do 4º aviado for inferior a 150% do 2º prescrito
                                    a4p2.nivel = 2
                                    Select Case genlab(p2Array(7), a4array(7), p2Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 4)
                                            anular(2)
                                            a4p2.nivel = 0
                                        Case 2
                                            descodificar(8, 4)
                                            a4p2.nivel = 2
                                        Case 3
                                            descodificar(7, 4)
                                            a4p2.nivel = 2
                                        Case 4
                                            result4.Text = "ver se há autorização (de " & p2Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Yellow
                                            a4p2.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 4º aviado NÃO for inferior a 150% do 2º prescrito
                                    Select Case genlab(p2Array(7), a4array(7), p2Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 4)
                                            anular(2)
                                            a4p2.nivel = 2 'nivel do h normal
                                        Case 2
                                            result4.Text = "H) + Aviamento de marca (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 3
                                            result4.Text = "H) + marca (" & p2Array(0) & ") trocado para genérico (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 4
                                            result4.Text = "H) + ver se há autorização (de " & p2Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Red
                                            a4p2.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a4p2.nivel = 3
                                    End Select
                                    a4p2.mostrado = False
                                    a4p1.mostrado = False
                                    a4p3.mostrado = False
                                    a4p4.mostrado = False
                                    a4p2.resultado = 2
                                    a4p2.nivel = 2.5
                                End If
                            Else 'se a quantidade do 4º aviado NÃO contiver só algarismos
                                a4p2.resultado = 1
                                a4p2.nivel = 3
                                descodificar(1, 4)
                            End If
                            av4.nivel = a4p2.nivel
                        Else 'se a dosagem do 4º aviado NÃO for igual ao do 2º prescrito
                            a4p2.resultado = 3
                            a4p2.nivel = 4
                            descodificar(3, 4)
                        End If
                    Else 'se a apresentação do 4º aviado NÃO for igual ao do 2º prescrito
                        a4p2.resultado = 4
                        a4p2.nivel = 5
                        'descodificar(4, 4)
                    End If
                Else 'se o dci do 4º aviado NÃO for igual ao do 2º prescrito
                    a4p2.resultado = 5
                    a4p2.nivel = 6
                    'descodificar(5, 4)
                End If
            End If
        Else
            If IsNothing(p2row) Then
                prescNexist(2)
            End If
            If IsNothing(a4row) Then
                aviadoNexist(4)
                msgNexist()
            End If
        End If
        If result4.Text = "" Then
            nPresc(4)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p3a1()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) And Not IsNothing(p3row) Then 'se o código do 1º aviado e 3º prescrito existem
            Dim a1array = a1row.ItemArray
            Dim p3Array = p3row.ItemArray
            If a1array(0).ToString = p3Array(0).ToString Then 'se o código do 1º aviado for igual ao do 3º prescrito
                a1p3.nivel = 0
                a1p3.resultado = 0
                descodificar(0, 1)
            Else
                If a1array(1).ToString = p3Array(1).ToString Then 'se o dci do 1º aviado for igual ao do 3º prescrito
                    If Via(a1array(2).ToString) = Via(p3Array(2).ToString) Then 'se a apresentação do 1º aviado for igual ao do 3º prescrito
                        If a1array(3).ToString = p3Array(3).ToString Then 'se a dosagem do 1º aviado for igual ao do 3º prescrito
                            If a1array(4).ToString = p3Array(4).ToString Then
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                Select Case genlab(p3Array(7), a1array(7), p3Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 1)
                                        anular(3)
                                        a1p3.nivel = 0
                                    Case 2
                                        descodificar(8, 1)
                                        a1p3.nivel = 2
                                    Case 3
                                        descodificar(7, 1)
                                        a1p3.nivel = 2
                                    Case 4
                                        result1.Text = "ver se há autorização (de " & p3Array(8) & " para " & a1array(8) & ")"
                                        result1.BackColor = Color.Yellow
                                        a1p3.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p3Array(4).ToString Then 'se a quantidade do 1º aviado contiver só algarismos
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                If a1array(4).ToString <= 1.5 * p3Array(4).ToString Then 'se a quantidade do 1º aviado for inferior a 150% do 3º prescrito
                                    a1p3.nivel = 2
                                    Select Case genlab(p3Array(7), a1array(7), p3Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 1)
                                            anular(3)
                                            a1p3.nivel = 0
                                        Case 2
                                            descodificar(8, 1)
                                            a1p3.nivel = 2
                                        Case 3
                                            descodificar(7, 1)
                                            a1p3.nivel = 2
                                        Case 4
                                            result1.Text = "ver se há autorização (de " & p3Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Yellow
                                            a1p3.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 1º aviado NÃO for inferior a 150% do 3º prescrito
                                    Select Case genlab(p3Array(7), a1array(7), p3Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 1)
                                            anular(3)
                                            a1p3.nivel = 2 'nivel do h normal
                                        Case 2
                                            result1.Text = "H) + Aviamento de marca (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 3
                                            result1.Text = "H) + marca (" & p3Array(0) & ") trocado para genérico (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 4
                                            result1.Text = "H) + ver se há autorização (de " & p3Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Red
                                            a1p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a1p3.nivel = 3
                                    End Select
                                    a1p2.mostrado = False
                                    a1p1.mostrado = False
                                    a1p3.mostrado = False
                                    a1p4.mostrado = False
                                    a1p3.resultado = 2
                                    a1p3.nivel = 2.5
                                End If
                            Else 'se a quantidade do 1º aviado NÃO contiver só algarismos
                                a1p3.resultado = 1
                                a1p3.nivel = 3
                                descodificar(1, 1)
                            End If
                            av1.nivel = a1p3.nivel
                        Else 'se a dosagem do 1º aviado NÃO for igual ao do 3º prescrito
                            a1p3.resultado = 3
                            a1p3.nivel = 4
                            descodificar(3, 1)
                        End If
                    Else 'se a apresentação do 1º aviado NÃO for igual ao do 3º prescrito
                        a1p3.resultado = 4
                        a1p3.nivel = 5
                        'descodificar(4, 1)
                    End If
                Else 'se o dci do 1º aviado NÃO for igual ao do 3º prescrito
                    a1p3.resultado = 5
                    a1p3.nivel = 6
                    'descodificar(5, 1)
                End If
            End If

        Else
            If IsNothing(p3row) Then
                prescNexist(3)
            End If
            If IsNothing(a1row) Then
                aviadoNexist(1)
                msgNexist()
            End If
        End If
        If result1.Text = "" Then
            nPresc(1)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p3a2()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) And Not IsNothing(p3row) Then 'se o código do 2º aviado e 3º prescrito existem
            Dim a2array = a2row.ItemArray
            Dim p3Array = p3row.ItemArray
            If a2array(0).ToString = p3Array(0).ToString Then 'se o código do 2º aviado for igual ao do 3º prescrito
                a2p3.nivel = 0
                a2p3.resultado = 0
                descodificar(0, 2)
            Else
                If a2array(1).ToString = p3Array(1).ToString Then 'se o dci do 2º aviado for igual ao do 3º prescrito
                    If Via(a2array(2).ToString) = Via(p3Array(2).ToString) Then 'se a apresentação do 2º aviado for igual ao do 3º prescrito
                        If a2array(3).ToString = p3Array(3).ToString Then 'se a dosagem do 2º aviado for igual ao do 3º prescrito
                            If a2array(4).ToString = p3Array(4).ToString Then
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                Select Case genlab(p3Array(7), a2array(7), p3Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 2)
                                        anular(3)
                                        a2p3.nivel = 0
                                    Case 2
                                        descodificar(8, 2)
                                        a2p3.nivel = 2
                                    Case 3
                                        descodificar(7, 2)
                                        a2p3.nivel = 2
                                    Case 4
                                        result2.Text = "ver se há autorização (de " & p3Array(8) & " para " & a2array(8) & ")"
                                        result2.BackColor = Color.Yellow
                                        a2p3.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p3Array(4).ToString Then 'se a quantidade do 2º aviado contiver só algarismos
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                If a2array(4).ToString <= 1.5 * p3Array(4).ToString Then 'se a quantidade do 2º aviado for inferior a 150% do 3º prescrito
                                    a2p3.nivel = 2
                                    Select Case genlab(p3Array(7), a2array(7), p3Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 2)
                                            anular(3)
                                            a2p3.nivel = 0
                                        Case 2
                                            descodificar(8, 2)
                                            a2p3.nivel = 2
                                        Case 3
                                            descodificar(7, 2)
                                            a2p3.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p3Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            a2p3.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 2º aviado NÃO for inferior a 150% do 3º prescrito
                                    Select Case genlab(p3Array(7), a2array(7), p3Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 2)
                                            anular(3)
                                            a2p3.nivel = 2 'nivel do h normal
                                        Case 2
                                            result2.Text = "H) + Aviamento de marca (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 3
                                            result2.Text = "H) + marca (" & p3Array(0) & ") trocado para genérico (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 4
                                            result2.Text = "H) + ver se há autorização (de " & p3Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Red
                                            a2p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a2p3.nivel = 3
                                    End Select
                                    a2p2.mostrado = False
                                    a2p1.mostrado = False
                                    a2p3.mostrado = False
                                    a2p4.mostrado = False
                                    a2p3.resultado = 2
                                    a2p3.nivel = 2.5
                                End If
                            Else 'se a quantidade do 2º aviado NÃO contiver só algarismos
                                a2p3.resultado = 1
                                a2p3.nivel = 3
                                descodificar(1, 2)
                            End If
                            av2.nivel = a2p3.nivel
                        Else 'se a dosagem do 2º aviado NÃO for igual ao do 3º prescrito
                            a2p3.resultado = 3
                            a2p3.nivel = 4
                            descodificar(3, 2)
                        End If
                    Else 'se a apresentação do 2º aviado NÃO for igual ao do 3º prescrito
                        a2p3.resultado = 4
                        a2p3.nivel = 5
                        'descodificar(4, 2)
                    End If
                Else 'se o dci do 2º aviado NÃO for igual ao do 3º prescrito
                    a2p3.resultado = 5
                    a2p3.nivel = 6
                    'descodificar(5, 2)
                End If
            End If

        Else
            If IsNothing(p3row) Then
                prescNexist(3)
            End If
            If IsNothing(a2row) Then
                aviadoNexist(2)
                msgNexist()
            End If
        End If
        If result2.Text = "" Then
            nPresc(2)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p3a3()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) And Not IsNothing(p3row) Then 'se o código do 3º aviado e 3º prescrito existem
            Dim a3array = a3row.ItemArray
            Dim p3Array = p3row.ItemArray
            If a3array(0).ToString = p3Array(0).ToString Then 'se o código do 3º aviado for igual ao do 3º prescrito
                a3p3.nivel = 0
                a3p3.resultado = 0
                descodificar(0, 3)
            Else
                If a3array(1).ToString = p3Array(1).ToString Then 'se o dci do 3º aviado for igual ao do 3º prescrito
                    If Via(a3array(2).ToString) = Via(p3Array(2).ToString) Then 'se a apresentação do 3º aviado for igual ao do 3º prescrito
                        If a3array(3).ToString = p3Array(3).ToString Then 'se a dosagem do 3º aviado for igual ao do 3º prescrito
                            If a3array(4).ToString = p3Array(4).ToString Then
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                Select Case genlab(p3Array(7), a3array(7), p3Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 3)
                                        anular(3)
                                        a3p3.nivel = 0
                                    Case 2
                                        descodificar(8, 3)
                                        a3p3.nivel = 2
                                    Case 3
                                        descodificar(7, 3)
                                        a3p3.nivel = 2
                                    Case 4
                                        result3.Text = "ver se há autorização (de " & p3Array(8) & " para " & a3array(8) & ")"
                                        result3.BackColor = Color.Yellow
                                        a3p3.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p3Array(4).ToString Then 'se a quantidade do 3º aviado contiver só algarismos
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                If a3array(4).ToString <= 1.5 * p3Array(4).ToString Then 'se a quantidade do 3º aviado for inferior a 150% do 3º prescrito
                                    a3p3.nivel = 2
                                    Select Case genlab(p3Array(7), a3array(7), p3Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 3)
                                            anular(3)
                                            a3p3.nivel = 0
                                        Case 2
                                            descodificar(8, 3)
                                            a3p3.nivel = 2
                                        Case 3
                                            descodificar(7, 3)
                                            a3p3.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p3Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            a3p3.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 3º aviado NÃO for inferior a 150% do 3º prescrito
                                    Select Case genlab(p3Array(7), a3array(7), p3Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 3)
                                            anular(3)
                                            a3p3.nivel = 2 'nivel do h normal
                                        Case 2
                                            result3.Text = "H) + Aviamento de marca (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 3
                                            result3.Text = "H) + marca (" & p3Array(0) & ") trocado para genérico (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 4
                                            result3.Text = "H) + ver se há autorização (de " & p3Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Red
                                            a3p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a3p3.nivel = 3
                                    End Select
                                    a3p2.mostrado = False
                                    a3p1.mostrado = False
                                    a3p3.mostrado = False
                                    a3p4.mostrado = False
                                    a3p3.resultado = 2
                                    a3p3.nivel = 2.5
                                End If
                            Else 'se a quantidade do 3º aviado NÃO contiver só algarismos
                                a3p3.resultado = 1
                                a3p3.nivel = 3
                                descodificar(1, 3)
                            End If
                            av3.nivel = a3p3.nivel
                        Else 'se a dosagem do 3º aviado NÃO for igual ao do 3º prescrito
                            a3p3.resultado = 3
                            a3p3.nivel = 4
                            descodificar(3, 3)
                        End If
                    Else 'se a apresentação do 3º aviado NÃO for igual ao do 3º prescrito
                        a3p3.resultado = 4
                        a3p3.nivel = 5
                        'descodificar(4, 3)
                    End If
                Else 'se o dci do 3º aviado NÃO for igual ao do 3º prescrito
                    a3p3.resultado = 5
                    a3p3.nivel = 6
                    'descodificar(5, 3)
                End If
            End If

        Else
            If IsNothing(p3row) Then
                prescNexist(3)
            End If
            If IsNothing(a3row) Then
                aviadoNexist(3)
                msgNexist()
            End If
        End If
        If result3.Text = "" Then
            nPresc(3)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p3a4()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) And Not IsNothing(p3row) Then 'se o código do 4º aviado e 3º prescrito existem
            Dim a4array = a4row.ItemArray
            Dim p3Array = p3row.ItemArray
            If a4array(0).ToString = p3Array(0).ToString Then 'se o código do 4º aviado for igual ao do 3º prescrito
                a4p3.nivel = 0
                a4p3.resultado = 0
                descodificar(0, 4)
            Else
                If a4array(1).ToString = p3Array(1).ToString Then 'se o dci do 4º aviado for igual ao do 3º prescrito
                    If Via(a4array(2).ToString) = Via(p3Array(2).ToString) Then 'se a apresentação do 4º aviado for igual ao do 3º prescrito
                        If a4array(3).ToString = p3Array(3).ToString Then 'se a dosagem do 4º aviado for igual ao do 3º prescrito
                            If a4array(4).ToString = p3Array(4).ToString Then
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                Select Case genlab(p3Array(7), a4array(7), p3Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 4)
                                        anular(3)
                                        a4p3.nivel = 0
                                    Case 2
                                        descodificar(8, 4)
                                        a4p3.nivel = 2
                                    Case 3
                                        descodificar(7, 4)
                                        a4p3.nivel = 2
                                    Case 4
                                        result4.Text = "ver se há autorização (de " & p3Array(8) & " para " & a4array(8) & ")"
                                        result4.BackColor = Color.Yellow
                                        a4p3.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a4array(4).ToString) And Not a4array(4).ToString = p3Array(4).ToString Then 'se a quantidade do 4º aviado contiver só algarismos
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                If a4array(4).ToString <= 1.5 * p3Array(4).ToString Then 'se a quantidade do 4º aviado for inferior a 150% do 3º prescrito
                                    a4p3.nivel = 2
                                    Select Case genlab(p3Array(7), a4array(7), p3Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 4)
                                            anular(3)
                                            a4p3.nivel = 0
                                        Case 2
                                            descodificar(8, 4)
                                            a4p3.nivel = 2
                                        Case 3
                                            descodificar(7, 4)
                                            a4p3.nivel = 2
                                        Case 4
                                            result4.Text = "ver se há autorização (de " & p3Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Yellow
                                            a4p3.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 4º aviado NÃO for inferior a 150% do 3º prescrito
                                    Select Case genlab(p3Array(7), a4array(7), p3Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 4)
                                            anular(3)
                                            a4p3.nivel = 2 'nivel do h normal
                                        Case 2
                                            result4.Text = "H) + Aviamento de marca (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 3
                                            result4.Text = "H) + marca (" & p3Array(0) & ") trocado para genérico (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 4
                                            result4.Text = "H) + ver se há autorização (de " & p3Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Red
                                            a4p3.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a4p3.nivel = 3
                                    End Select
                                    a4p2.mostrado = False
                                    a4p1.mostrado = False
                                    a4p3.mostrado = False
                                    a4p4.mostrado = False
                                    a4p3.resultado = 2
                                    a4p3.nivel = 2.5
                                End If
                            Else 'se a quantidade do 4º aviado NÃO contiver só algarismos
                                a4p3.resultado = 1
                                a4p3.nivel = 3
                                descodificar(1, 4)
                            End If
                            av4.nivel = a4p3.nivel
                        Else 'se a dosagem do 4º aviado NÃO for igual ao do 3º prescrito
                            a4p3.resultado = 3
                            a4p3.nivel = 4
                            descodificar(3, 4)
                        End If
                    Else 'se a apresentação do 4º aviado NÃO for igual ao do 3º prescrito
                        a4p3.resultado = 4
                        a4p3.nivel = 5
                        'descodificar(4, 4)
                    End If
                Else 'se o dci do 4º aviado NÃO for igual ao do 3º prescrito
                    a4p3.resultado = 5
                    a4p3.nivel = 6
                    'descodificar(5, 4)
                End If
            End If

        Else
            If IsNothing(p3row) Then
                prescNexist(3)
            End If
            If IsNothing(a4row) Then
                aviadoNexist(4)
                msgNexist()
            End If
        End If
        If result4.Text = "" Then
            nPresc(4)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p4a1()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a1row) And Not IsNothing(p4row) Then 'se o código do 1º aviado e 4º prescrito existem
            Dim a1array = a1row.ItemArray
            Dim p4Array = p4row.ItemArray
            If a1array(0).ToString = p4Array(0).ToString Then 'se o código do 1º aviado for igual ao do 4º prescrito
                a1p4.nivel = 0
                a1p4.resultado = 0
                descodificar(0, 1)
            Else
                If a1array(1).ToString = p4Array(1).ToString Then 'se o dci do 1º aviado for igual ao do 4º prescrito
                    If Via(a1array(2).ToString) = Via(p4Array(2).ToString) Then 'se a apresentação do 1º aviado for igual ao do 4º prescrito
                        If a1array(3).ToString = p4Array(3).ToString Then 'se a dosagem do 1º aviado for igual ao do 4º prescrito
                            If a1array(4).ToString = p4Array(4).ToString Then
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                Select Case genlab(p4Array(7), a1array(7), p4Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 1)
                                        anular(4)
                                        a1p4.nivel = 0
                                    Case 2
                                        descodificar(8, 1)
                                        a1p4.nivel = 2
                                    Case 3
                                        descodificar(7, 1)
                                        a1p4.nivel = 2
                                    Case 4
                                        result1.Text = "ver se há autorização (de " & p4Array(8) & " para " & a1array(8) & ")"
                                        result1.BackColor = Color.Yellow
                                        a1p4.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a1array(4).ToString) And Not a1array(4).ToString = p4Array(4).ToString Then 'se a quantidade do 1º aviado contiver só algarismos
                                a1p2.mostrado = True
                                a1p1.mostrado = True
                                a1p3.mostrado = True
                                a1p4.mostrado = True
                                If a1array(4).ToString <= 1.5 * p4Array(4).ToString Then 'se a quantidade do 1º aviado for inferior a 150% do 4º prescrito
                                    a1p4.nivel = 2
                                    Select Case genlab(p4Array(7), a1array(7), p4Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 1)
                                            anular(4)
                                            a1p4.nivel = 0
                                        Case 2
                                            descodificar(8, 1)
                                            a1p4.nivel = 2
                                        Case 3
                                            descodificar(7, 1)
                                            a1p4.nivel = 2
                                        Case 4
                                            result1.Text = "ver se há autorização (de " & p4Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Yellow
                                            a1p4.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 1º aviado NÃO for inferior a 150% do 4º prescrito
                                    Select Case genlab(p4Array(7), a1array(7), p4Array(8), a1array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 1)
                                            anular(4)
                                            a1p4.nivel = 2 'nivel do h normal
                                        Case 2
                                            result1.Text = "H) + Aviamento de marca (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 3
                                            result1.Text = "H) + marca (" & p4Array(0) & ") trocado para genérico (" & a1array(0) & ")"
                                            result1.BackColor = Color.Red
                                            a1p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case 4
                                            result1.Text = "H) + ver se há autorização (de " & p4Array(8) & " para " & a1array(8) & ")"
                                            result1.BackColor = Color.Red
                                            a1p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a1array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a1p4.nivel = 3
                                    End Select
                                    a1p2.mostrado = False
                                    a1p1.mostrado = False
                                    a1p3.mostrado = False
                                    a1p4.mostrado = False
                                    a1p4.resultado = 2
                                    a1p4.nivel = 2.5
                                End If
                            Else 'se a quantidade do 1º aviado NÃO contiver só algarismos
                                a1p4.resultado = 1
                                a1p4.nivel = 3
                                descodificar(1, 1)
                            End If
                            av1.nivel = a1p4.nivel
                        Else 'se a dosagem do 1º aviado NÃO for igual ao do 4º prescrito
                            a1p4.resultado = 3
                            a1p4.nivel = 4
                            descodificar(3, 1)
                        End If
                    Else 'se a apresentação do 1º aviado NÃO for igual ao do 4º prescrito
                        a1p4.resultado = 4
                        a1p4.nivel = 5
                        'descodificar(4, 1)
                    End If
                Else 'se o dci do 1º aviado NÃO for igual ao do 4º prescrito
                    a1p4.resultado = 5
                    a1p4.nivel = 6
                    'descodificar(5, 1)
                End If
            End If
        Else
            If IsNothing(p4row) Then
                prescNexist(4)
            End If
            If IsNothing(a1row) Then
                aviadoNexist(1)
                msgNexist()
            End If
        End If
        If result1.Text = "" Then
            nPresc(1)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p4a2()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a2row) And Not IsNothing(p4row) Then 'se o código do 2º aviado e 4º prescrito existem
            Dim a2array = a2row.ItemArray
            Dim p4Array = p4row.ItemArray
            If a2array(0).ToString = p4Array(0).ToString Then 'se o código do 2º aviado for igual ao do 4º prescrito
                a2p4.nivel = 0
                a2p4.resultado = 0
                descodificar(0, 2)
            Else
                If a2array(1).ToString = p4Array(1).ToString Then 'se o dci do 2º aviado for igual ao do 4º prescrito
                    If Via(a2array(2).ToString) = Via(p4Array(2).ToString) Then 'se a apresentação do 2º aviado for igual ao do 4º prescrito
                        If a2array(3).ToString = p4Array(3).ToString Then 'se a dosagem do 2º aviado for igual ao do 4º prescrito
                            If a2array(4).ToString = p4Array(4).ToString Then
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                Select Case genlab(p4Array(7), a2array(7), p4Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 2)
                                        anular(4)
                                        a2p4.nivel = 0
                                    Case 2
                                        descodificar(8, 2)
                                        a2p4.nivel = 2
                                    Case 3
                                        descodificar(7, 2)
                                        a2p4.nivel = 2
                                    Case 4
                                        result2.Text = "ver se há autorização (de " & p4Array(8) & " para " & a2array(8) & ")"
                                        result2.BackColor = Color.Yellow
                                        a2p4.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a2array(4).ToString) And Not a2array(4).ToString = p4Array(4).ToString Then 'se a quantidade do 2º aviado contiver só algarismos
                                a2p2.mostrado = True
                                a2p1.mostrado = True
                                a2p3.mostrado = True
                                a2p4.mostrado = True
                                If a2array(4).ToString <= 1.5 * p4Array(4).ToString Then 'se a quantidade do 2º aviado for inferior a 150% do 4º prescrito
                                    a2p4.nivel = 2
                                    Select Case genlab(p4Array(7), a2array(7), p4Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 2)
                                            anular(4)
                                            a2p4.nivel = 0
                                        Case 2
                                            descodificar(8, 2)
                                            a2p4.nivel = 2
                                        Case 3
                                            descodificar(7, 2)
                                            a2p4.nivel = 2
                                        Case 4
                                            result2.Text = "ver se há autorização (de " & p4Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Yellow
                                            a2p4.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 2º aviado NÃO for inferior a 150% do 4º prescrito
                                    Select Case genlab(p4Array(7), a2array(7), p4Array(8), a2array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 2)
                                            anular(4)
                                            a2p4.nivel = 2 'nivel do h normal
                                        Case 2
                                            result2.Text = "H) + Aviamento de marca (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 3
                                            result2.Text = "H) + marca (" & p4Array(0) & ") trocado para genérico (" & a2array(0) & ")"
                                            result2.BackColor = Color.Red
                                            a2p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case 4
                                            result2.Text = "H) + ver se há autorização (de " & p4Array(8) & " para " & a2array(8) & ")"
                                            result2.BackColor = Color.Red
                                            a2p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a2array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a2p4.nivel = 3
                                    End Select
                                    a2p2.mostrado = False
                                    a2p1.mostrado = False
                                    a2p3.mostrado = False
                                    a2p4.mostrado = False
                                    a2p4.resultado = 2
                                    a2p4.nivel = 2.5
                                End If
                            Else 'se a quantidade do 2º aviado NÃO contiver só algarismos
                                a2p4.resultado = 1
                                a2p4.nivel = 3
                                descodificar(1, 2)
                            End If
                            av2.nivel = a2p4.nivel
                        Else 'se a dosagem do 2º aviado NÃO for igual ao do 4º prescrito
                            a2p4.resultado = 3
                            a2p4.nivel = 4
                            descodificar(3, 2)
                        End If
                    Else 'se a apresentação do 2º aviado NÃO for igual ao do 4º prescrito
                        a2p4.resultado = 4
                        a2p4.nivel = 5
                        'descodificar(4, 2)
                    End If
                Else 'se o dci do 2º aviado NÃO for igual ao do 4º prescrito
                    a2p4.resultado = 5
                    a2p4.nivel = 6
                    'descodificar(5, 2)
                End If
            End If

        Else
            If IsNothing(p4row) Then
                prescNexist(4)
            End If
            If IsNothing(a2row) Then
                aviadoNexist(2)
                msgNexist()
            End If
        End If
        If result2.Text = "" Then
            nPresc(2)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p4a3()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a3row) And Not IsNothing(p4row) Then 'se o código do 3º aviado e 4º prescrito existem
            Dim a3array = a3row.ItemArray
            Dim p4Array = p4row.ItemArray
            If a3array(0).ToString = p4Array(0).ToString Then 'se o código do 3º aviado for igual ao do 4º prescrito
                a3p4.nivel = 0
                a3p4.resultado = 0
                descodificar(0, 3)
            Else
                If a3array(1).ToString = p4Array(1).ToString Then 'se o dci do 3º aviado for igual ao do 4º prescrito
                    If Via(a3array(2).ToString) = Via(p4Array(2).ToString) Then 'se a apresentação do 3º aviado for igual ao do 4º prescrito
                        If a3array(3).ToString = p4Array(3).ToString Then 'se a dosagem do 3º aviado for igual ao do 4º prescrito
                            If a3array(4).ToString = p4Array(4).ToString Then
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                Select Case genlab(p4Array(7), a3array(7), p4Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 3)
                                        anular(4)
                                        a3p4.nivel = 0
                                    Case 2
                                        descodificar(8, 3)
                                        a3p4.nivel = 2
                                    Case 3
                                        descodificar(7, 3)
                                        a3p4.nivel = 2
                                    Case 4
                                        result3.Text = "ver se há autorização (de " & p4Array(8) & " para " & a3array(8) & ")"
                                        result3.BackColor = Color.Yellow
                                        a3p4.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a3array(4).ToString) And Not a3array(4).ToString = p4Array(4).ToString Then 'se a quantidade do 3º aviado contiver só algarismos
                                a3p2.mostrado = True
                                a3p1.mostrado = True
                                a3p3.mostrado = True
                                a3p4.mostrado = True
                                If a3array(4).ToString <= 1.5 * p4Array(4).ToString Then 'se a quantidade do 3º aviado for inferior a 150% do 4º prescrito
                                    a3p4.nivel = 2
                                    Select Case genlab(p4Array(7), a3array(7), p4Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 3)
                                            anular(4)
                                            a3p4.nivel = 0
                                        Case 2
                                            descodificar(8, 3)
                                            a3p4.nivel = 2
                                        Case 3
                                            descodificar(7, 3)
                                            a3p4.nivel = 2
                                        Case 4
                                            result3.Text = "ver se há autorização (de " & p4Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Yellow
                                            a3p4.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 3º aviado NÃO for inferior a 150% do 4º prescrito
                                    Select Case genlab(p4Array(7), a3array(7), p4Array(8), a3array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 3)
                                            anular(4)
                                            a3p4.nivel = 2 'nivel do h normal
                                        Case 2
                                            result3.Text = "H) + Aviamento de marca (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 3
                                            result3.Text = "H) + marca (" & p4Array(0) & ") trocado para genérico (" & a3array(0) & ")"
                                            result3.BackColor = Color.Red
                                            a3p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case 4
                                            result3.Text = "H) + ver se há autorização (de " & p4Array(8) & " para " & a3array(8) & ")"
                                            result3.BackColor = Color.Red
                                            a3p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a3array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a3p4.nivel = 3
                                    End Select
                                    a3p2.mostrado = False
                                    a3p1.mostrado = False
                                    a3p3.mostrado = False
                                    a3p4.mostrado = False
                                    a3p4.resultado = 2
                                    a3p4.nivel = 2.5
                                End If
                            Else 'se a quantidade do 3º aviado NÃO contiver só algarismos
                                a3p4.resultado = 1
                                a3p4.nivel = 3
                                descodificar(1, 3)
                            End If
                            av3.nivel = a3p4.nivel
                        Else 'se a dosagem do 3º aviado NÃO for igual ao do 4º prescrito
                            a3p4.resultado = 3
                            a3p4.nivel = 4
                            descodificar(3, 3)
                        End If
                    Else 'se a apresentação do 3º aviado NÃO for igual ao do 4º prescrito
                        a3p4.resultado = 4
                        a3p4.nivel = 5
                        'descodificar(4, 3)
                    End If
                Else 'se o dci do 3º aviado NÃO for igual ao do 4º prescrito
                    a3p4.resultado = 5
                    a3p4.nivel = 6
                    'descodificar(5, 3)
                End If
            End If

        Else
            If IsNothing(p4row) Then
                prescNexist(4)
            End If
            If IsNothing(a3row) Then
                aviadoNexist(3)
                msgNexist()
            End If
        End If
        If result3.Text = "" Then
            nPresc(3)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub p4a4()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(a4row) And Not IsNothing(p4row) Then 'se o código do 4º aviado e 4º prescrito existem
            Dim a4array = a4row.ItemArray
            Dim p4Array = p4row.ItemArray
            If a4array(0).ToString = p4Array(0).ToString Then 'se o código do 4º aviado for igual ao do 4º prescrito
                a4p4.nivel = 0
                a4p4.resultado = 0
                descodificar(0, 4)
            Else
                If a4array(1).ToString = p4Array(1).ToString Then 'se o dci do 4º aviado for igual ao do 4º prescrito
                    If Via(a4array(2).ToString) = Via(p4Array(2).ToString) Then 'se a apresentação do 4º aviado for igual ao do 4º prescrito
                        If a4array(3).ToString = p4Array(3).ToString Then 'se a dosagem do 4º aviado for igual ao do 4º prescrito
                            If a4array(4).ToString = p4Array(4).ToString Then
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                Select Case genlab(p4Array(7), a4array(7), p4Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                    Case 1
                                        descodificar(0, 4)
                                        anular(4)
                                        a4p4.nivel = 0
                                    Case 2
                                        descodificar(8, 4)
                                        a4p4.nivel = 2
                                    Case 3
                                        descodificar(7, 4)
                                        a4p4.nivel = 2
                                    Case 4
                                        result4.Text = "ver se há autorização (de " & p4Array(8) & " para " & a4array(8) & ")"
                                        result4.BackColor = Color.Yellow
                                        a4p4.nivel = 2
                                    Case Else
                                        MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                End Select
                            ElseIf IsNumeric(a4array(4).ToString) And Not a4array(4).ToString = p4Array(4).ToString Then 'se a quantidade do 4º aviado contiver só algarismos
                                a4p2.mostrado = True
                                a4p1.mostrado = True
                                a4p3.mostrado = True
                                a4p4.mostrado = True
                                If a4array(4).ToString <= 1.5 * p4Array(4).ToString Then 'se a quantidade do 4º aviado for inferior a 150% do 4º prescrito
                                    a4p4.nivel = 2
                                    Select Case genlab(p4Array(7), a4array(7), p4Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(0, 4)
                                            anular(4)
                                            a4p4.nivel = 0
                                        Case 2
                                            descodificar(8, 4)
                                            a4p4.nivel = 2
                                        Case 3
                                            descodificar(7, 4)
                                            a4p4.nivel = 2
                                        Case 4
                                            result4.Text = "ver se há autorização (de " & p4Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Yellow
                                            a4p4.nivel = 2
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                    End Select
                                Else 'se a quantidade do 4º aviado NÃO for inferior a 150% do 4º prescrito
                                    Select Case genlab(p4Array(7), a4array(7), p4Array(8), a4array(8)) 'gen->gen, marca->gen, gen->marca
                                        Case 1
                                            descodificar(2, 4)
                                            anular(4)
                                            a4p4.nivel = 2 'nivel do h normal
                                        Case 2
                                            result4.Text = "H) + Aviamento de marca (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 3
                                            result4.Text = "H) + marca (" & p4Array(0) & ") trocado para genérico (" & a4array(0) & ")"
                                            result4.BackColor = Color.Red
                                            a4p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case 4
                                            result4.Text = "H) + ver se há autorização (de " & p4Array(8) & " para " & a4array(8) & ")"
                                            result4.BackColor = Color.Red
                                            a4p4.nivel = 2 'nivel daqui
                                            MsgBox("troca no " & a4array(0))
                                        Case Else
                                            MsgBox("genlab não devolve comparação", MsgBoxStyle.OkOnly)
                                            a4p4.nivel = 3
                                    End Select
                                    a4p2.mostrado = False
                                    a4p1.mostrado = False
                                    a4p3.mostrado = False
                                    a4p4.mostrado = False
                                    a4p4.resultado = 2
                                    a4p4.nivel = 2.5
                                End If
                            Else 'se a quantidade do 4º aviado NÃO contiver só algarismos
                                a4p4.resultado = 1
                                a4p4.nivel = 3
                                descodificar(1, 4)
                            End If
                            av4.nivel = a4p4.nivel
                        Else 'se a dosagem do 4º aviado NÃO for igual ao do 4º prescrito
                            a4p4.resultado = 3
                            a4p4.nivel = 4
                            descodificar(3, 4)
                        End If
                    Else 'se a apresentação do 4º aviado NÃO for igual ao do 4º prescrito
                        a4p4.resultado = 4
                        a4p4.nivel = 5
                        'descodificar(4, 4)
                    End If
                Else 'se o dci do 4º aviado NÃO for igual ao do 4º prescrito
                    a4p4.resultado = 5
                    a4p4.nivel = 6
                    'descodificar(5, 4)
                End If
            End If

        Else
            If IsNothing(p4row) Then
                prescNexist(4)
            End If
            If IsNothing(a4row) Then
                aviadoNexist(4)
                msgNexist()
            End If
        End If
        If result4.Text = "" Then
            nPresc(4)
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Function genlab(ByVal pgen As Boolean, ByVal agen As Boolean, ByVal plab As String, ByVal alab As String) As Short
        On Error GoTo MOSTRARERRO
1:      Select Case agen
            Case False
3:              If pgen = False Then
4:                  Return 1
5:              Else
6:                  Return 2
7:              End If
8:          Case True
9:              If pgen = False Then
10:                 Return 3
11:             Else
12:                 If alab = plab Then
13:                     Return 1
14:                 Else
15:                     Return 4
16:                 End If
17:             End If
18:             End Select
        Exit Function
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Sub OK(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "OK"
4:              result1.BackColor = Color.Green
5:          Case 2
6:              result2.Text = "OK"
7:              result2.BackColor = Color.Green
8:          Case 3
9:              result3.Text = "OK"
10:             result3.BackColor = Color.Green
11:         Case 4
12:             result4.Text = "OK"
13:             result4.BackColor = Color.Green
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub Hquant(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "H) Quantidade do aviado excede em mais de 50% o prescrito"
4:              result1.BackColor = Color.Red
5:          Case 2
6:              result2.Text = "H) Quantidade do aviado excede em mais de 50% o prescrito"
7:              result2.BackColor = Color.Red
8:          Case 3
9:              result3.Text = "H) Quantidade do aviado excede em mais de 50% o prescrito"
10:             result3.BackColor = Color.Red
11:         Case 4
12:             result4.Text = "H) Quantidade do aviado excede em mais de 50% o prescrito"
13:             result4.BackColor = Color.Red
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub dciDif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.BackColor = Color.Red
4:              result1.Text = "G) DCI não prescrito"
5:          Case 2
6:              result2.BackColor = Color.Red
7:              result2.Text = "G) DCI não prescrito"
8:          Case 3
9:              result3.BackColor = Color.Red
10:             result3.Text = "G) DCI não prescrito"
11:         Case 4
12:             result4.BackColor = Color.Red
13:             result4.Text = "G) DCI não prescrito"
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub apresDif(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.BackColor = Color.Red
4:              result1.Text = "G) apresentação diferente da prescrita"
5:          Case 2
6:              result2.BackColor = Color.Red
7:              result2.Text = "G) apresentação diferente da prescrita"
8:          Case 3
9:              result3.BackColor = Color.Red
10:             result3.Text = "G) apresentação diferente da prescrita"
11:         Case 4
12:             result4.BackColor = Color.Red
13:             result4.Text = "G) apresentação diferente da prescrita"
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Function anular(ByVal qual As Short) As Boolean
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              procurado1 = True
4:          Case 2
5:              procurado2 = True
6:          Case 3
7:              procurado3 = True
8:          Case 4
9:              procurado4 = True
10:             End Select
        Exit Function
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Sub nPresc(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      If A > P Then
2:          Select Case qual
                Case 1
4:                  result1.Text = "G) embalagem não prescrita"
5:                  result1.BackColor = Color.Red
6:              Case 2
7:                  result2.Text = "G) embalagem não prescrita"
8:                  result2.BackColor = Color.Red
9:              Case 3
10:                 result3.Text = "G) embalagem não prescrita"
11:                 result3.BackColor = Color.Red
12:             Case 4
13:                 result4.Text = "G) embalagem não prescrita"
14:                 result4.BackColor = Color.Red
15:                 End Select
16:     Else
17:         Select Case qual
                Case 1
19:                 result1.Text = "G) DCI não prescrito"
20:                 result1.BackColor = Color.Red
21:             Case 2
22:                 result2.Text = "G) DCI não prescrito"
23:                 result2.BackColor = Color.Red
24:             Case 3
25:                 result3.Text = "G) DCI não prescrito"
26:                 result3.BackColor = Color.Red
27:             Case 4
28:                 result4.Text = "G) DCI não prescrito"
29:                 result4.BackColor = Color.Red
30:                 End Select
31:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub marca(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "Aviado de marca!!!"
4:              result1.BackColor = Color.Orange
5:          Case 2
6:              result2.Text = "Aviado de marca!!!"
7:              result2.BackColor = Color.Orange
8:          Case 3
9:              result3.Text = "Aviado de marca!!!"
10:             result3.BackColor = Color.Orange
11:         Case 4
12:             result4.Text = "Aviado de marca!!!"
13:             result4.BackColor = Color.Orange
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub marca2gen(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "ver se há autorização [Marca -> Genérico]"
4:              result1.BackColor = Color.Yellow
5:          Case 2
6:              result2.Text = "ver se há autorização [Marca -> Genérico]"
7:              result2.BackColor = Color.Yellow
8:          Case 3
9:              result3.Text = "ver se há autorização [Marca -> Genérico]"
10:             result3.BackColor = Color.Yellow
11:         Case 4
12:             result4.Text = "ver se há autorização [Marca -> Genérico]"
13:             result4.BackColor = Color.Yellow
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub nComp(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              result1.Text = "F) embalagem não comparticipada"
4:              result1.BackColor = Color.Red
5:          Case 2
6:              result2.Text = "F) embalagem não comparticipada"
7:              result2.BackColor = Color.Red
8:          Case 3
9:              result3.Text = "F) embalagem não comparticipada"
10:             result3.BackColor = Color.Red
11:         Case 4
12:             result4.Text = "F) embalagem não comparticipada"
13:             result4.BackColor = Color.Red
14:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
7:          Case 2
8:              If aviam2.Text <> "0" And aviam2.Text <> "C" And aviam2.Text <> "C0" Then
9:                  result2.BackColor = Color.Gray
10:                 result2.Text = "o código aviado não existe"
11:             End If
12:         Case 3
13:             If aviam3.Text <> "0" And aviam3.Text <> "C" And aviam3.Text <> "C0" Then
14:                 result3.BackColor = Color.Gray
15:                 result3.Text = "o código aviado não existe"
16:             End If
17:         Case 4
18:             If aviam4.Text <> "0" And aviam4.Text <> "C" And aviam4.Text <> "C0" Then
19:                 result4.BackColor = Color.Gray
20:                 result4.Text = "o código aviado não existe"
21:             End If
22:         Case 9
23:             If codEC.Text <> "0" And codEC.Text <> "C" And codEC.Text <> "C0" Then
24:                 labelmedcompgenports.BackColor = Color.Gray
25:                 labelmedcompgenports.Text = "o código aviado não existe"
26:             End If
27:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub prescNexist(ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case qual
            Case 1
3:              If semcod1.Checked = False And presc1.Text <> "0" Then
4:                  presc1.BackColor = Color.Red
5:              End If
6:          Case 2
7:              If semcod2.Checked = False And presc2.Text <> "0" Then
8:                  presc2.BackColor = Color.Red
9:              End If
10:         Case 3
11:             If semcod3.Checked = False And presc3.Text <> "0" Then
12:                 presc3.BackColor = Color.Red
13:             End If
14:         Case 4
15:             If semcod4.Checked = False And presc4.Text <> "0" Then
16:                 presc4.BackColor = Color.Red
17:             End If
18:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub fazer99(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
1:      qual.nivel = 99
2:      qual.resultado = 99
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub







    Sub agrupar()
        On Error GoTo MOSTRARERRO
        a1p1.grupo = 1
        a2p1.grupo = 1
        a3p1.grupo = 1
        a4p1.grupo = 1
        grupoP1 = 1
        If Not IsNothing(p1row) Then
            grupoP1dci = p1row(1).ToString
        End If
        grupoA1 = 1
        If Not IsNothing(a1row) Then
            grupoA1dci = a1row(1).ToString
        End If
        If IsNothing(p1row) And semcod1.Checked = False Then
            presc1.BackColor = Color.Red
        End If
        If P >= 2 Then
            If Not IsNothing(p2row) Then
                Dim p2array = p2row.ItemArray
                If Not IsNothing(p1row) Then
                    Dim p1array = p1row.ItemArray
                    Select Case p2array(1).ToString
                        Case p1array(1).ToString
                            a1p2.grupo = 1
                            a2p2.grupo = 1
                            a3p2.grupo = 1
                            a4p2.grupo = 1
                            grupoP1 = grupoP1 + 1
                            grupoP2dci = p2array(1).ToString
                            presc1.BackColor = Color.Yellow
                            presc2.BackColor = Color.Yellow
                        Case Else
                            a1p2.grupo = 2
                            a2p2.grupo = 2
                            a3p2.grupo = 2
                            a4p2.grupo = 2
                            grupoP2 = grupoP2 + 1
                            grupoP2dci = p2array(1).ToString
                    End Select
                Else
                    prescNexist(1)
                End If
            Else
                prescNexist(2)
            End If

            If P >= 3 Then
                If Not IsNothing(p3row) Then
                    Dim p3array = p3row.ItemArray
                    If Not IsNothing(p1row) Then
                        Dim p1array = p1row.ItemArray
                        Select Case p3array(1).ToString
                            Case p1array(1).ToString
                                a1p3.grupo = 1
                                a2p3.grupo = 1
                                a3p3.grupo = 1
                                a4p3.grupo = 1
                                grupoP1 = grupoP1 + 1
                                grupoP1dci = p1array(1).ToString
                                presc1.BackColor = Color.Yellow
                                presc3.BackColor = Color.Yellow
                            Case Else
                                If Not IsNothing(p2row) Then
                                    Dim p2array = p2row.ItemArray
                                    If p3array(1).ToString = p2array(1).ToString Then
                                        a1p3.grupo = 2
                                        a2p3.grupo = 2
                                        a3p3.grupo = 2
                                        a4p3.grupo = 2
                                        grupoP2 = grupoP2 + 1
                                        grupoP2dci = p2array(1).ToString
                                        presc3.BackColor = Color.Yellow
                                        presc2.BackColor = Color.Yellow
                                    Else
                                        a1p3.grupo = 3
                                        a2p3.grupo = 3
                                        a3p3.grupo = 3
                                        a4p3.grupo = 3
                                        grupoP3 = grupoP3 + 1
                                        grupoP3dci = p3array(1).ToString
                                    End If
                                Else
                                    prescNexist(2)
                                End If
                        End Select
                    Else
                        prescNexist(1)
                    End If
                Else
                    prescNexist(3)
                End If
            End If
            If P >= 4 Then
                If Not IsNothing(p4row) Then
                    Dim p4array = p4row.ItemArray
                    If Not IsNothing(p1row) Then
                        Dim p1array = p1row.ItemArray
                        Select Case p4array(1).ToString
                            Case p1array(1).ToString
                                a1p4.grupo = 1
                                a2p4.grupo = 1
                                a3p4.grupo = 1
                                a4p4.grupo = 1
                                grupoP1 = grupoP1 + 1
                                grupoP1dci = p1array(1).ToString
                                presc1.BackColor = Color.Yellow
                                presc4.BackColor = Color.Yellow
                            Case Else
                                If Not IsNothing(p2row) Then
                                    Dim p2array = p2row.ItemArray
                                    If p4array(1).ToString = p2array(1).ToString Then
                                        a1p4.grupo = 2
                                        a2p4.grupo = 2
                                        a3p4.grupo = 2
                                        a4p4.grupo = 2
                                        grupoP2 = grupoP2 + 1
                                        grupoP2dci = p2array(1).ToString
                                        presc4.BackColor = Color.Yellow
                                        presc2.BackColor = Color.Yellow
                                    Else
                                        If Not IsNothing(p3row) Then
                                            Dim p3array = p3row.ItemArray
                                            If p4array(1).ToString = p3array(1).ToString Then
                                                a1p4.grupo = 3
                                                a2p4.grupo = 3
                                                a3p4.grupo = 3
                                                a4p4.grupo = 3
                                                grupoP3 = grupoP3 + 1
                                                grupoP3dci = p3array(1).ToString
                                                presc4.BackColor = Color.Yellow
                                                presc3.BackColor = Color.Yellow
                                            Else
                                                a1p4.grupo = 4
                                                a2p4.grupo = 4
                                                a3p4.grupo = 4
                                                a4p4.grupo = 4
                                                grupoP4 = grupoP4 + 1
                                                grupoP4dci = p4array(1).ToString
                                            End If
                                        Else
                                            prescNexist(3)
                                        End If
                                    End If
                                Else
                                    prescNexist(2)
                                End If
                        End Select
                    Else
                        prescNexist(1)
                    End If
                Else
                    prescNexist(4)
                End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub AgruparAviados()
        On Error GoTo MOSTRARERRO
        If A = 2 Then
            If Not IsNothing(a2row) Then
                Dim a2array = a2row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a2array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            grupoA2 = grupoA2 + 1
                            grupoA2dci = a2array(1).ToString
                    End Select
                End If
            End If
        End If
        If A = 3 Then
            If Not IsNothing(a2row) Then
                Dim a2array = a2row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a2array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            grupoA2 = grupoA2 + 1
                            grupoA2dci = a2array(1).ToString
                    End Select
                End If
            End If
            If Not IsNothing(a3row) Then
                Dim a3array = a3row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a3array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            If Not IsNothing(a2row) Then
                                Dim a2array = a2row.ItemArray
                                If a3array(1).ToString = a2array(1).ToString Then
                                    grupoA2 = grupoA2 + 1
                                    grupoA2dci = a2array(1).ToString
                                Else
                                    grupoA3 = grupoA3 + 1
                                    grupoA3dci = a3array(1).ToString
                                End If
                            End If
                    End Select
                End If
            End If
        End If
        If A = 4 Then
            If Not IsNothing(a2row) Then
                Dim a2array = a2row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a2array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            grupoA2 = grupoA2 + 1
                            grupoA2dci = a2array(1).ToString
                    End Select
                End If
            End If
            If Not IsNothing(a3row) Then
                Dim a3array = a3row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a3array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            If Not IsNothing(a2row) Then
                                Dim a2array = a2row.ItemArray
                                If a3array(1).ToString = a2array(1).ToString Then
                                    grupoA2 = grupoA2 + 1
                                    grupoA2dci = a2array(1).ToString
                                Else
                                    grupoA3 = grupoA3 + 1
                                    grupoA3dci = a3array(1).ToString
                                End If
                            End If
                    End Select
                End If
            End If
            If Not IsNothing(a4row) Then
                Dim a4array = a4row.ItemArray
                If Not IsNothing(a1row) Then
                    Dim a1array = a1row.ItemArray
                    Select Case a4array(1).ToString
                        Case a1array(1).ToString
                            grupoA1 = grupoA1 + 1
                            grupoA1dci = a1array(1).ToString
                        Case Else
                            If Not IsNothing(a2row) Then
                                Dim a2array = a2row.ItemArray
                                If a4array(1).ToString = a2array(1).ToString Then
                                    grupoA2 = grupoA2 + 1
                                    grupoA2dci = a2array(1).ToString
                                Else
                                    If Not IsNothing(a3row) Then
                                        Dim a3array = a3row.ItemArray
                                        If a4array(1).ToString = a3array(1).ToString Then
                                            grupoA3 = grupoA3 + 1
                                            grupoA3dci = a3array(1).ToString
                                        Else
                                            grupoA4 = grupoA4 + 1
                                            grupoA4dci = a4array(1).ToString
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                End If
            End If
        End If


        If grupoA1 >= 2 Then
            If grupoA1dci = grupoP1dci And grupoA1 > grupoP1 Then
                If A >= 2 And Not IsNothing(a2row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If grupoA1dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p1.nivel < a3p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If
            If grupoA1dci = grupoP2dci And grupoA1 > grupoP2 Then
                If grupoA1dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If grupoA1dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If
            If grupoA1dci = grupoP3dci And grupoA1 > grupoP3 Then
                If A >= 2 And Not IsNothing(a2row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 2 And Not IsNothing(a2row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If grupoA1dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If
            If grupoA1dci = grupoP4dci And grupoA1 > grupoP4 Then
                If A >= 2 And Not IsNothing(a2row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 2 And Not IsNothing(a2row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA1dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If
        End If

        If grupoA2 >= 2 Then
            If grupoA2dci = grupoP1dci And grupoA2 > grupoP1 Then
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p1.nivel < a1p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If
            If grupoA2dci = grupoP2dci And grupoA2 > grupoP2 Then
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 3 And Not IsNothing(a3row) Then
                    If grupoA2dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                        descodificar(5, 4)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                        descodificar(5, 1)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                        descodificar(5, 2)
                    End If
                End If
                If A >= 4 And Not IsNothing(a4row) Then
                    If grupoA2dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                        descodificar(5, 3)
                    End If
                End If
            End If

            If grupoA2dci = grupoP3dci And grupoA2 > grupoP3 Then
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p3.nivel < a3p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA2dci = grupoP4dci And grupoA2 > grupoP4 Then
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA2dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA2dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA2dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
        End If

        If grupoA3 >= 2 Then
            If grupoA3dci = grupoP1dci And grupoA3 > grupoP1 Then
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA3dci = grupoP2dci And grupoA3 > grupoP2 Then
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA3dci = grupoP3dci And grupoA3 > grupoP3 Then
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA3dci = grupoP4dci And grupoA3 > grupoP4 Then
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA3dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA3dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
        End If

        If grupoA4 >= 2 Then

            If grupoA4dci = grupoP1dci And grupoA1 > grupoP1 Then
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p1.nivel < a4p1.nivel And a4p1.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p1.nivel < a1p1.nivel And a1p1.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p1.nivel < a2p1.nivel And a2p1.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p1.nivel < a3p1.nivel And a3p1.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA4dci = grupoP2dci And grupoA4 > grupoP2 Then
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p2.nivel < a4p2.nivel And a4p2.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p2.nivel < a1p2.nivel And a1p2.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p2.nivel < a2p2.nivel And a2p2.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p2.nivel < a3p2.nivel And a3p2.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA4dci = grupoP3dci And grupoA4 > grupoP3 Then
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p3.nivel < a4p3.nivel And a4p3.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p3.nivel < a1p3.nivel And a1p3.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p3.nivel < a2p3.nivel And a2p3.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p3.nivel < a3p3.nivel And a3p3.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
            If grupoA4dci = grupoP4dci And grupoA4 > grupoP4 Then
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a2row(1) And a1p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a3row(1) And a1p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a1row(1) And a1row(1) = a4row(1) And a1p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a1row(1) And a2p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a3row(1) And a2p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
                If grupoA4dci.ToString = a2row(1) And a2row(1) = a4row(1) And a2p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a1row(1) And a3p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a2row(1) And a3p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a3row(1) And a3row(1) = a4row(1) And a3p4.nivel < a4p4.nivel And a4p4.nivel <= 6 Then
                    descodificar(5, 4)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a1row(1) And a4p4.nivel < a1p4.nivel And a1p4.nivel <= 6 Then
                    descodificar(5, 1)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a2row(1) And a4p4.nivel < a2p4.nivel And a2p4.nivel <= 6 Then
                    descodificar(5, 2)
                End If
                If grupoA4dci.ToString = a4row(1) And a4row(1) = a3row(1) And a4p4.nivel < a3p4.nivel And a3p4.nivel <= 6 Then
                    descodificar(5, 3)
                End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub






    Sub prioridade()
        On Error GoTo MOSTRARERRO
        agrupar()
        If P = 1 Then
        End If
        If P = 2 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a2p2)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If

                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then

                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)

                        Else

                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a1p2)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)

                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then

                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)

                        Else

                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)

                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel Then

                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)

                        Else

                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)

                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)
                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        Else
                            descodificar((a4p2.resultado), 4)
                            'anular(2)
                            fazer99(a4p1)
                            IndicarComoMostrado(2, 4)
                            ApagarRepetido(2, 4)
                        End If
                    End If
                End If
            End If
        End If
        If P = 3 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            If a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel Then
                                descodificar((a3p2.resultado), 3)
                                'anular(2)
                                fazer99(a3p1)
                                fazer99(a3p3)
                                IndicarComoMostrado(2, 3)
                                ApagarRepetido(2, 3)
                            Else
                                descodificar((a3p3.resultado), 3)
                                'anular(3)
                                fazer99(a3p1)
                                fazer99(a3p2)
                                IndicarComoMostrado(3, 3)
                                ApagarRepetido(3, 3)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            If a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel Then
                                descodificar((a3p2.resultado), 3)
                                'anular(2)
                                fazer99(a3p1)
                                fazer99(a3p3)
                                IndicarComoMostrado(2, 3)
                                ApagarRepetido(2, 3)
                            Else
                                descodificar((a3p3.resultado), 3)
                                'anular(3)
                                fazer99(a3p1)
                                fazer99(a3p2)
                                IndicarComoMostrado(3, 3)
                                ApagarRepetido(3, 3)
                            End If
                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        Else
                            If a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel Then
                                descodificar((a4p2.resultado), 4)
                                ' anular(2)
                                fazer99(a4p1)
                                fazer99(a4p3)
                                IndicarComoMostrado(2, 4)
                                ApagarRepetido(2, 4)
                            Else
                                descodificar((a4p3.resultado), 4)
                                'anular(3)
                                fazer99(a4p1)
                                fazer99(a4p2)
                                IndicarComoMostrado(3, 4)
                                ApagarRepetido(3, 4)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If P = 4 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        ElseIf a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        ElseIf a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                            descodificar((a1p3.resultado), 1)
                            'anular(3)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p4)
                            IndicarComoMostrado(3, 1)
                            ApagarRepetido(3, 1)
                        Else
                            descodificar((a1p4.resultado), 1)
                            'anular(4)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(4, 1)
                            ApagarRepetido(4, 1)

                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        ElseIf a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        ElseIf a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                            descodificar((a1p3.resultado), 1)
                            ' anular(3)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p4)
                            IndicarComoMostrado(3, 1)
                            ApagarRepetido(3, 1)
                        Else
                            descodificar((a1p4.resultado), 1)
                            'anular(4)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(4, 1)
                            ApagarRepetido(4, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        ElseIf a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)
                        ElseIf a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p3.nivel And a2p3.nivel <= a2p4.nivel Then
                            descodificar((a2p3.resultado), 2)
                            'anular(3)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p4)
                            IndicarComoMostrado(3, 2)
                            ApagarRepetido(3, 2)
                        Else
                            descodificar((a2p4.resultado), 2)
                            'anular(4)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(4, 2)
                            ApagarRepetido(4, 2)

                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a3p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        ElseIf a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        ElseIf a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                            descodificar((a1p3.resultado), 1)
                            'anular(3)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p4)
                            IndicarComoMostrado(3, 1)
                            ApagarRepetido(3, 1)
                        Else
                            descodificar((a1p4.resultado), 1)
                            'anular(4)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(4, 1)
                            ApagarRepetido(4, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        ElseIf a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)
                        ElseIf a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p2.nivel And a2p3.nivel <= a2p4.nivel Then
                            descodificar((a2p3.resultado), 2)
                            'anular(3)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p4)
                            IndicarComoMostrado(3, 2)
                            ApagarRepetido(3, 2)
                        Else
                            descodificar((a2p4.resultado), 2)
                            'anular(4)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(4, 2)
                            ApagarRepetido(4, 2)

                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a1p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a1p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    ElseIf a3p4.resultado = 0 Then
                        OK(3)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a1p4)
                        fazer99(a4p4)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        IndicarComoMostrado(4, 3)
                        ApagarRepetido(4, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel And a3p1.nivel <= a3p4.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        ElseIf a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel And a3p2.nivel <= a3p4.nivel Then
                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)
                        ElseIf a3p3.nivel <= a3p1.nivel And a3p3.nivel <= a3p2.nivel And a3p3.nivel <= a3p4.nivel Then
                            descodificar((a3p3.resultado), 3)
                            'anular(3)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p4)
                            IndicarComoMostrado(3, 3)
                            ApagarRepetido(3, 3)
                        Else
                            descodificar((a3p4.resultado), 3)
                            'anular(4)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(4, 3)
                            ApagarRepetido(4, 3)

                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    ElseIf a4p4.resultado = 0 Then
                        OK(4)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        IndicarComoMostrado(4, 4)
                        ApagarRepetido(4, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel And a4p1.nivel <= a4p4.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            fazer99(a4p4)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        ElseIf a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel And a4p2.nivel <= a4p4.nivel Then
                            descodificar((a4p2.resultado), 4)
                            'anular(2)
                            fazer99(a4p1)
                            fazer99(a4p3)
                            fazer99(a4p4)
                            IndicarComoMostrado(2, 4)
                            ApagarRepetido(2, 4)
                        ElseIf a4p3.nivel <= a4p1.nivel And a4p3.nivel <= a4p2.nivel And a4p3.nivel <= a4p4.nivel Then
                            descodificar((a4p3.resultado), 4)
                            'anular(3)
                            fazer99(a4p1)
                            fazer99(a4p2)
                            fazer99(a4p4)
                            IndicarComoMostrado(3, 4)
                            ApagarRepetido(3, 4)
                        Else
                            descodificar((a4p4.resultado), 4)
                            'anular(4)
                            fazer99(a4p1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            IndicarComoMostrado(4, 4)
                            ApagarRepetido(4, 4)

                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p2)
                        fazer99(a1p1)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p1)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        ElseIf a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        ElseIf a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                            descodificar((a1p3.resultado), 1)
                            'anular(3)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p4)
                            IndicarComoMostrado(3, 1)
                            ApagarRepetido(3, 1)
                        Else
                            descodificar((a1p4.resultado), 1)
                            'anular(4)
                            fazer99(a1p1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(4, 1)
                            ApagarRepetido(4, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        ElseIf a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)
                        ElseIf a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p2.nivel And a2p3.nivel <= a2p4.nivel Then
                            descodificar((a2p3.resultado), 2)
                            'anular(3)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p4)
                            IndicarComoMostrado(3, 2)
                            ApagarRepetido(3, 2)
                        Else
                            descodificar((a2p4.resultado), 2)
                            'anular(4)
                            fazer99(a2p1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(4, 2)
                            ApagarRepetido(4, 2)

                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a1p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a1p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    ElseIf a3p4.resultado = 0 Then
                        OK(3)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a1p4)
                        fazer99(a4p4)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        IndicarComoMostrado(4, 3)
                        ApagarRepetido(4, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel And a3p1.nivel <= a3p4.nivel Then

                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)

                        ElseIf a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel And a3p2.nivel <= a3p4.nivel Then
                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)

                        ElseIf a3p3.nivel <= a3p1.nivel And a3p3.nivel <= a3p2.nivel And a3p3.nivel <= a3p4.nivel Then

                            descodificar((a3p3.resultado), 3)
                            'anular(3)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p4)
                            IndicarComoMostrado(3, 3)
                            ApagarRepetido(3, 3)

                        Else
                            descodificar((a3p4.resultado), 3)
                            'anular(4)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(4, 3)
                            ApagarRepetido(4, 3)



                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        'anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        'anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a2p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        'anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a2p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    ElseIf a4p4.resultado = 0 Then
                        OK(4)
                        'anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a2p4)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        IndicarComoMostrado(4, 4)
                        ApagarRepetido(4, 4)
                    Else

                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel And a4p1.nivel <= a4p4.nivel Then

                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            fazer99(a4p4)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)

                        ElseIf a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel And a4p2.nivel <= a4p4.nivel Then

                            descodificar((a4p2.resultado), 4)
                            'anular(2)
                            fazer99(a4p1)
                            fazer99(a4p3)
                            fazer99(a4p4)
                            IndicarComoMostrado(2, 4)
                            ApagarRepetido(2, 4)
                        ElseIf a4p3.nivel <= a4p1.nivel And a4p3.nivel <= a4p2.nivel And a4p3.nivel <= a4p4.nivel Then
                            descodificar((a4p3.resultado), 4)
                            'anular(3)
                            fazer99(a4p1)
                            fazer99(a4p2)
                            fazer99(a4p4)
                            IndicarComoMostrado(3, 4)
                            ApagarRepetido(3, 4)
                        Else
                            descodificar((a4p4.resultado), 4)
                            'anular(4)
                            fazer99(a4p1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            IndicarComoMostrado(4, 4)
                            ApagarRepetido(4, 4)

                        End If
                    End If

                End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    'diferente do outro prioridade pois inclui comparação de nivel entre A's dentro da comparação entre P's (incompleto - só até meio do p=2 a=4)
    'tem erro. não se pode mostrar um aX por comparação por aY só porque aY não foi mostrado e sem saber se aX foi ou não... altera já mostrados e tudo...
    Sub prioridade2()
        On Error GoTo MOSTRARERRO
        agrupar()
        If P = 1 Then
        End If
        If P = 2 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)
                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a2p2)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If

                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then

                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)

                        Else

                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a1p2)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)

                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then

                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)

                        Else

                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)

                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel Then

                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)

                        Else

                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)

                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        'fazer 99 a todos os aX e pY que não o aXpY
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    Else
                        'comparar entre p's
                        If a1p1.nivel <= a1p2.nivel Then

                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)

                        Else

                            descodificar((a1p2.resultado), 1)
                            'anular(2)
                            fazer99(a1p1)
                            IndicarComoMostrado(2, 1)
                            ApagarRepetido(2, 1)

                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            descodificar((a2p2.resultado), 2)
                            'anular(2)
                            fazer99(a2p1)
                            IndicarComoMostrado(2, 2)
                            ApagarRepetido(2, 2)
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)
                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        Else
                            descodificar((a4p2.resultado), 4)
                            'anular(2)
                            fazer99(a4p1)
                            IndicarComoMostrado(2, 4)
                            ApagarRepetido(2, 4)
                        End If
                    End If
                End If
            End If
        End If
        If P = 3 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            If a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel Then
                                descodificar((a3p2.resultado), 3)
                                'anular(2)
                                fazer99(a3p1)
                                fazer99(a3p3)
                                IndicarComoMostrado(2, 3)
                                ApagarRepetido(2, 3)
                            Else
                                descodificar((a3p3.resultado), 3)
                                'anular(3)
                                fazer99(a3p1)
                                fazer99(a3p2)
                                IndicarComoMostrado(3, 3)
                                ApagarRepetido(3, 3)
                            End If
                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                descodificar((a1p3.resultado), 1)
                                'anular(3)
                                fazer99(a1p1)
                                fazer99(a1p2)
                                IndicarComoMostrado(3, 1)
                                ApagarRepetido(3, 1)
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                descodificar((a2p3.resultado), 2)
                                'anular(3)
                                fazer99(a2p1)
                                fazer99(a2p2)
                                IndicarComoMostrado(3, 2)
                                ApagarRepetido(3, 2)
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            If a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel Then
                                descodificar((a3p2.resultado), 3)
                                'anular(2)
                                fazer99(a3p1)
                                fazer99(a3p3)
                                IndicarComoMostrado(2, 3)
                                ApagarRepetido(2, 3)
                            Else
                                descodificar((a3p3.resultado), 3)
                                'anular(3)
                                fazer99(a3p1)
                                fazer99(a3p2)
                                IndicarComoMostrado(3, 3)
                                ApagarRepetido(3, 3)
                            End If
                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        Else
                            If a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel Then
                                descodificar((a4p2.resultado), 4)
                                ' anular(2)
                                fazer99(a4p1)
                                fazer99(a4p3)
                                IndicarComoMostrado(2, 4)
                                ApagarRepetido(2, 4)
                            Else
                                descodificar((a4p3.resultado), 4)
                                'anular(3)
                                fazer99(a4p1)
                                fazer99(a4p2)
                                IndicarComoMostrado(3, 4)
                                ApagarRepetido(3, 4)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If P = 4 Then
            'relembrar que quanto menor o nível maior a prioridade(0 a 9)
            'os que não receberam nivel por não exitirem têem nivel de 99 atribuido na inicialização
            If A = 1 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                fazer99(a1p4)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                If a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                                    descodificar((a1p3.resultado), 1)
                                    'anular(3)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p4)
                                    IndicarComoMostrado(3, 1)
                                    ApagarRepetido(3, 1)
                                Else
                                    descodificar((a1p4.resultado), 1)
                                    'anular(4)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p3)
                                    IndicarComoMostrado(4, 1)
                                    ApagarRepetido(4, 1)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If A = 2 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                fazer99(a1p4)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                If a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                                    descodificar((a1p3.resultado), 1)
                                    ' anular(3)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p4)
                                    IndicarComoMostrado(3, 1)
                                    ApagarRepetido(3, 1)
                                Else
                                    descodificar((a1p4.resultado), 1)
                                    'anular(4)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p3)
                                    IndicarComoMostrado(4, 1)
                                    ApagarRepetido(4, 1)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                fazer99(a2p4)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                If a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p3.nivel And a2p3.nivel <= a2p4.nivel Then
                                    descodificar((a2p3.resultado), 2)
                                    'anular(3)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p4)
                                    IndicarComoMostrado(3, 2)
                                    ApagarRepetido(3, 2)
                                Else
                                    descodificar((a2p4.resultado), 2)
                                    'anular(4)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p3)
                                    IndicarComoMostrado(4, 2)
                                    ApagarRepetido(4, 2)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If A = 3 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a3p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                fazer99(a1p4)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                If a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                                    descodificar((a1p3.resultado), 1)
                                    'anular(3)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p4)
                                    IndicarComoMostrado(3, 1)
                                    ApagarRepetido(3, 1)
                                Else
                                    descodificar((a1p4.resultado), 1)
                                    'anular(4)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p3)
                                    IndicarComoMostrado(4, 1)
                                    ApagarRepetido(4, 1)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                fazer99(a2p4)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                If a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p2.nivel And a2p3.nivel <= a2p4.nivel Then
                                    descodificar((a2p3.resultado), 2)
                                    'anular(3)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p4)
                                    IndicarComoMostrado(3, 2)
                                    ApagarRepetido(3, 2)
                                Else
                                    descodificar((a2p4.resultado), 2)
                                    'anular(4)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p3)
                                    IndicarComoMostrado(4, 2)
                                    ApagarRepetido(4, 2)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a1p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a1p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    ElseIf a3p4.resultado = 0 Then
                        OK(3)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a1p4)
                        fazer99(a4p4)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        IndicarComoMostrado(4, 3)
                        ApagarRepetido(4, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel And a3p1.nivel <= a3p4.nivel Then
                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)
                        Else
                            If a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel And a3p2.nivel <= a3p4.nivel Then
                                descodificar((a3p2.resultado), 3)
                                'anular(2)
                                fazer99(a3p1)
                                fazer99(a3p3)
                                fazer99(a3p4)
                                IndicarComoMostrado(2, 3)
                                ApagarRepetido(2, 3)
                            Else
                                If a3p3.nivel <= a3p1.nivel And a3p3.nivel <= a3p2.nivel And a3p3.nivel <= a3p4.nivel Then
                                    descodificar((a3p3.resultado), 3)
                                    'anular(3)
                                    fazer99(a3p1)
                                    fazer99(a3p2)
                                    fazer99(a3p4)
                                    IndicarComoMostrado(3, 3)
                                    ApagarRepetido(3, 3)
                                Else
                                    descodificar((a3p4.resultado), 3)
                                    'anular(4)
                                    fazer99(a3p1)
                                    fazer99(a3p2)
                                    fazer99(a3p3)
                                    IndicarComoMostrado(4, 3)
                                    ApagarRepetido(4, 3)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    ElseIf a4p4.resultado = 0 Then
                        OK(4)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        IndicarComoMostrado(4, 4)
                        ApagarRepetido(4, 4)
                    Else
                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel And a4p1.nivel <= a4p4.nivel Then
                            descodificar((a4p1.resultado), 4)
                            'anular(1)
                            fazer99(a4p2)
                            fazer99(a4p3)
                            fazer99(a4p4)
                            IndicarComoMostrado(1, 4)
                            ApagarRepetido(1, 4)
                        Else
                            If a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel And a4p2.nivel <= a4p4.nivel Then
                                descodificar((a4p2.resultado), 4)
                                'anular(2)
                                fazer99(a4p1)
                                fazer99(a4p3)
                                fazer99(a4p4)
                                IndicarComoMostrado(2, 4)
                                ApagarRepetido(2, 4)
                            Else
                                If a4p3.nivel <= a4p1.nivel And a4p3.nivel <= a4p2.nivel And a4p3.nivel <= a4p4.nivel Then
                                    descodificar((a4p3.resultado), 4)
                                    'anular(3)
                                    fazer99(a4p1)
                                    fazer99(a4p2)
                                    fazer99(a4p4)
                                    IndicarComoMostrado(3, 4)
                                    ApagarRepetido(3, 4)
                                Else
                                    descodificar((a4p4.resultado), 4)
                                    'anular(4)
                                    fazer99(a4p1)
                                    fazer99(a4p2)
                                    fazer99(a4p3)
                                    IndicarComoMostrado(4, 4)
                                    ApagarRepetido(4, 4)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If A = 4 Then
                If Not a1p2.mostrado = True Then
                    If a1p1.resultado = 0 Then
                        OK(1)
                        anular(1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(1, 1)
                        ApagarRepetido(1, 1)
                    ElseIf a1p2.resultado = 0 Then
                        OK(1)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a1p1)
                        fazer99(a1p3)
                        fazer99(a1p4)
                        IndicarComoMostrado(2, 1)
                        ApagarRepetido(2, 1)
                    ElseIf a1p3.resultado = 0 Then
                        OK(1)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a1p2)
                        fazer99(a1p1)
                        fazer99(a1p4)
                        IndicarComoMostrado(3, 1)
                        ApagarRepetido(3, 1)
                    ElseIf a1p4.resultado = 0 Then
                        OK(1)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a1p2)
                        fazer99(a1p3)
                        fazer99(a1p1)
                        IndicarComoMostrado(4, 1)
                        ApagarRepetido(4, 1)
                    Else
                        If a1p1.nivel <= a1p2.nivel And a1p1.nivel <= a1p3.nivel And a1p1.nivel <= a1p4.nivel Then
                            descodificar((a1p1.resultado), 1)
                            'anular(1)
                            fazer99(a1p2)
                            fazer99(a1p3)
                            fazer99(a1p4)
                            IndicarComoMostrado(1, 1)
                            ApagarRepetido(1, 1)
                        Else
                            If a1p2.nivel <= a1p1.nivel And a1p2.nivel <= a1p3.nivel And a1p2.nivel <= a1p4.nivel Then
                                descodificar((a1p2.resultado), 1)
                                'anular(2)
                                fazer99(a1p1)
                                fazer99(a1p3)
                                fazer99(a1p4)
                                IndicarComoMostrado(2, 1)
                                ApagarRepetido(2, 1)
                            Else
                                If a1p3.nivel <= a1p1.nivel And a1p3.nivel <= a1p2.nivel And a1p3.nivel <= a1p4.nivel Then
                                    descodificar((a1p3.resultado), 1)
                                    'anular(3)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p4)
                                    IndicarComoMostrado(3, 1)
                                    ApagarRepetido(3, 1)
                                Else
                                    descodificar((a1p4.resultado), 1)
                                    'anular(4)
                                    fazer99(a1p1)
                                    fazer99(a1p2)
                                    fazer99(a1p3)
                                    IndicarComoMostrado(4, 1)
                                    ApagarRepetido(4, 1)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a2p2.mostrado = True Then
                    If a2p1.resultado = 0 Then
                        OK(2)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a3p1)
                        fazer99(a4p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(1, 2)
                        ApagarRepetido(1, 2)
                    ElseIf a2p2.resultado = 0 Then
                        OK(2)
                        anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a4p2)
                        fazer99(a2p1)
                        fazer99(a2p3)
                        fazer99(a2p4)
                        IndicarComoMostrado(2, 2)
                        ApagarRepetido(2, 2)
                    ElseIf a2p3.resultado = 0 Then
                        OK(2)
                        anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a4p3)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p4)
                        IndicarComoMostrado(3, 2)
                        ApagarRepetido(3, 2)
                    ElseIf a2p4.resultado = 0 Then
                        OK(2)
                        anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a4p4)
                        fazer99(a2p1)
                        fazer99(a2p2)
                        fazer99(a2p3)
                        IndicarComoMostrado(4, 2)
                        ApagarRepetido(4, 2)
                    Else
                        If a2p1.nivel <= a2p2.nivel And a2p1.nivel <= a2p3.nivel And a2p1.nivel <= a2p4.nivel Then
                            descodificar((a2p1.resultado), 2)
                            'anular(1)
                            fazer99(a2p2)
                            fazer99(a2p3)
                            fazer99(a2p4)
                            IndicarComoMostrado(1, 2)
                            ApagarRepetido(1, 2)
                        Else
                            If a2p2.nivel <= a2p1.nivel And a2p2.nivel <= a2p3.nivel And a2p2.nivel <= a2p4.nivel Then
                                descodificar((a2p2.resultado), 2)
                                'anular(2)
                                fazer99(a2p1)
                                fazer99(a2p3)
                                fazer99(a2p4)
                                IndicarComoMostrado(2, 2)
                                ApagarRepetido(2, 2)
                            Else
                                If a2p3.nivel <= a2p1.nivel And a2p3.nivel <= a2p2.nivel And a2p3.nivel <= a2p4.nivel Then
                                    descodificar((a2p3.resultado), 2)
                                    'anular(3)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p4)
                                    IndicarComoMostrado(3, 2)
                                    ApagarRepetido(3, 2)
                                Else
                                    descodificar((a2p4.resultado), 2)
                                    'anular(4)
                                    fazer99(a2p1)
                                    fazer99(a2p2)
                                    fazer99(a2p3)
                                    IndicarComoMostrado(4, 2)
                                    ApagarRepetido(4, 2)
                                End If
                            End If
                        End If
                    End If
                End If
                If Not a3p2.mostrado = True Then
                    If a3p1.resultado = 0 Then
                        OK(3)
                        anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a4p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(1, 3)
                        ApagarRepetido(1, 3)
                    ElseIf a3p2.resultado = 0 Then
                        OK(3)
                        anular(2)
                        fazer99(a2p2)
                        fazer99(a1p2)
                        fazer99(a4p2)
                        fazer99(a3p1)
                        fazer99(a3p3)
                        fazer99(a3p4)
                        IndicarComoMostrado(2, 3)
                        ApagarRepetido(2, 3)
                    ElseIf a3p3.resultado = 0 Then
                        OK(3)
                        anular(3)
                        fazer99(a2p3)
                        fazer99(a1p3)
                        fazer99(a4p3)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p4)
                        IndicarComoMostrado(3, 3)
                        ApagarRepetido(3, 3)
                    ElseIf a3p4.resultado = 0 Then
                        OK(3)
                        anular(4)
                        fazer99(a2p4)
                        fazer99(a1p4)
                        fazer99(a4p4)
                        fazer99(a3p1)
                        fazer99(a3p2)
                        fazer99(a3p3)
                        IndicarComoMostrado(4, 3)
                        ApagarRepetido(4, 3)
                    Else
                        If a3p1.nivel <= a3p2.nivel And a3p1.nivel <= a3p3.nivel And a3p1.nivel <= a3p4.nivel Then


                            descodificar((a3p1.resultado), 3)
                            'anular(1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(1, 3)
                            ApagarRepetido(1, 3)

                        ElseIf a3p2.nivel <= a3p1.nivel And a3p2.nivel <= a3p3.nivel And a3p2.nivel <= a3p4.nivel Then
                            descodificar((a3p2.resultado), 3)
                            'anular(2)
                            fazer99(a3p1)
                            fazer99(a3p3)
                            fazer99(a3p4)
                            IndicarComoMostrado(2, 3)
                            ApagarRepetido(2, 3)
                        ElseIf a3p3.nivel <= a3p1.nivel And a3p3.nivel <= a3p2.nivel And a3p3.nivel <= a3p4.nivel Then

                            descodificar((a3p3.resultado), 3)
                            'anular(3)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p4)
                            IndicarComoMostrado(3, 3)
                            ApagarRepetido(3, 3)

                        Else

                            descodificar((a3p4.resultado), 3)
                            'anular(4)
                            fazer99(a3p1)
                            fazer99(a3p2)
                            fazer99(a3p3)
                            IndicarComoMostrado(4, 3)
                            ApagarRepetido(4, 3)


                        End If
                    End If
                End If
                If Not a4p2.mostrado = True Then
                    If a4p1.resultado = 0 Then
                        OK(4)
                        'anular(1)
                        fazer99(a1p1)
                        fazer99(a2p1)
                        fazer99(a3p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(1, 4)
                        ApagarRepetido(1, 4)
                    ElseIf a4p2.resultado = 0 Then
                        OK(4)
                        'anular(2)
                        fazer99(a1p2)
                        fazer99(a3p2)
                        fazer99(a2p2)
                        fazer99(a4p1)
                        fazer99(a4p3)
                        fazer99(a4p4)
                        IndicarComoMostrado(2, 4)
                        ApagarRepetido(2, 4)
                    ElseIf a4p3.resultado = 0 Then
                        OK(4)
                        'anular(3)
                        fazer99(a1p3)
                        fazer99(a3p3)
                        fazer99(a2p3)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p4)
                        IndicarComoMostrado(3, 4)
                        ApagarRepetido(3, 4)
                    ElseIf a4p4.resultado = 0 Then
                        OK(4)
                        'anular(4)
                        fazer99(a1p4)
                        fazer99(a3p4)
                        fazer99(a2p4)
                        fazer99(a4p1)
                        fazer99(a4p2)
                        fazer99(a4p3)
                        IndicarComoMostrado(4, 4)
                        ApagarRepetido(4, 4)
                    Else

                        If a4p1.nivel <= a4p2.nivel And a4p1.nivel <= a4p3.nivel And a4p1.nivel <= a4p4.nivel Then
                            If a3p1.nivel <= a1p1.nivel And a3p1.nivel <= a2p1.nivel And a3p1.nivel <= a4p1.nivel Then
                                'MsgBox("antes")
                                descodificar((a3p1.resultado), 3)
                                'anular(1)
                                fazer99(a3p2)
                                fazer99(a3p3)
                                fazer99(a3p4)
                                IndicarComoMostrado(1, 3)
                                ApagarRepetido(1, 3)
                                'MsgBox("depois")
                            ElseIf a4p1.nivel <= a1p1.nivel And a4p1.nivel <= a2p1.nivel And a4p1.nivel <= a3p1.nivel Then
                                descodificar((a4p1.resultado), 4)
                                'anular(1)
                                fazer99(a4p2)
                                fazer99(a4p3)
                                fazer99(a4p4)
                                IndicarComoMostrado(1, 4)
                                ApagarRepetido(1, 4)

                            Else

                            End If

                        Else
                            If a4p2.nivel <= a4p1.nivel And a4p2.nivel <= a4p3.nivel And a4p2.nivel <= a4p4.nivel Then
                                descodificar((a4p2.resultado), 4)
                                'anular(2)
                                fazer99(a4p1)
                                fazer99(a4p3)
                                fazer99(a4p4)
                                IndicarComoMostrado(2, 4)
                                ApagarRepetido(2, 4)
                            Else

                                If a4p3.nivel <= a4p1.nivel And a4p3.nivel <= a4p2.nivel And a4p3.nivel <= a4p4.nivel Then
                                    descodificar((a4p3.resultado), 4)
                                    'anular(3)
                                    fazer99(a4p1)
                                    fazer99(a4p2)
                                    fazer99(a4p4)
                                    IndicarComoMostrado(3, 4)
                                    ApagarRepetido(3, 4)
                                Else
                                    descodificar((a4p4.resultado), 4)
                                    'anular(4)
                                    fazer99(a4p1)
                                    fazer99(a4p2)
                                    fazer99(a4p3)
                                    IndicarComoMostrado(4, 4)
                                    ApagarRepetido(4, 4)
                                End If
                            End If
                        End If
                    End If

                End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub descodificar(ByVal resultado As Short, ByVal qual As Short)
        On Error GoTo MOSTRARERRO
1:      Select Case resultado
            Case 0
3:              OK(qual)
4:          Case 1
5:              verifQuant(qual)
6:          Case 2
7:              Hquant(qual)
8:          Case 3
9:              DoseDif(qual)
10:         Case 4
11:             apresDif(qual)
12:         Case 5
13:             dciDif(qual)
14:         Case 6
15:             nPresc(qual)
16:         Case 7
17:             marca2gen(qual)
18:         Case 8
19:             marca(qual)
20:         Case 9
21:             nComp(qual)
22:         Case 10
23:             aviadoNexist(qual)
24:         Case 99
25:             'MsgBox("nivel 99" & vbCr & "resultado=" & resultado & vbTab & qual)
26:             'não fazer nada - código de resultado previsto por isso não pode despolotar o else
27:         Case Else
28:             MsgBox("sub descodificar recebeu código de resultado inválido para A= " & qual)
29:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub








    'mudança no 1º checkbox - inutiliza caixa de texto e poe fundo cinza ou repõe
    Private Sub semcod1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles semcod1.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If semcod1.Checked = True Then
2:          presc1.Text = "0"
3:          presc1.BackColor = Color.LightGray
4:      Else
5:          presc1.BackColor = Color.White
6:      End If
7:      Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'mudança no 2º checkbox - inutiliza caixa de texto e poe fundo cinza ou repõe
    Private Sub semcod2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles semcod2.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If semcod1.Checked = True Or presc1.Text <> "0" Then
2:          If semcod2.Checked = True Then
3:              presc2.Text = "0"
4:              presc2.BackColor = Color.LightGray
5:          Else
6:              presc2.BackColor = Color.White
7:          End If
8:      Else
9:          Beep()
10:         semcod2.Checked = False
11:         Me.presc1.Focus()
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'mudança no 3º checkbox - inutiliza caixa de texto e poe fundo cinza ou repõe
    Private Sub semcod3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles semcod3.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If semcod2.Checked = True Or presc2.Text <> "0" Then
2:          If semcod3.Checked = True Then
3:              presc3.Text = "0"
4:              presc3.BackColor = Color.LightGray
5:          Else
6:              presc3.BackColor = Color.White
7:          End If
8:      Else
9:          Beep()
10:         semcod3.Checked = False
11:         Me.presc2.Focus()
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'mudança no 4º checkbox - inutiliza caixa de texto e poe fundo cinza ou repõe
    Private Sub semcod4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles semcod4.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If semcod3.Checked = True Or presc3.Text <> "0" Then
2:          If semcod4.Checked = True Then
3:              presc4.Text = "0"
4:              presc4.BackColor = Color.LightGray
5:          Else
6:              presc4.BackColor = Color.White
7:          End If
8:      Else
9:          Beep()
10:         semcod4.Checked = False
11:         Me.presc3.Focus()
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Dim linha1(0 To 8)
    Dim linha2(0 To 8)
    Dim linha3(0 To 8)
    Dim linha4(0 To 8)

    'a1array.Add(linha1) - colocar isto no if semcod1.checked=true
    'dá erro. corrigir e transpor para os outros aviam's e os outros subaviam's
    'em todas as verificações de prescX.text e de pXrow(0) e pXarray(0) tem de se colocar antes o if semcodX.checked=True

    Private Sub presc1dci_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1dci.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha1(1) = presc1dci.Text
        'p1row(1).ItemArray = presc1dci.Text
        'p1row(1) = presc1dci.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc1forma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1forma.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha1(2) = presc1forma.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc1dose_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1dose.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha1(3) = presc1dose.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc1qty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1qty.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha1(4) = presc1qty.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc1lab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1lab.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha1(5) = presc1lab.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc1gen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc1gen.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If presc1gen.Checked = True Then
2:          linha1(6) = "True"
3:      Else
4:          linha1(6) = "False"
5:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Private Sub presc2dci_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2dci.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha2(1) = presc2dci.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc2forma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2forma.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha2(2) = presc2forma.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc2dose_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2dose.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha2(3) = presc2dose.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc2qty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2qty.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha2(4) = presc2qty.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc2lab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2lab.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha2(8) = presc2lab.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc2gen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc2gen.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If presc2gen.Checked = True Then
2:          linha2(7) = "True"
3:      Else
4:          linha2(7) = "False"
5:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Private Sub presc3dci_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3dci.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha3(1) = presc3dci.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc3forma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3forma.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha3(2) = presc3forma.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc3dose_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3dose.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha3(3) = presc3dose.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc3qty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3qty.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha3(4) = presc3qty.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc3lab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3lab.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha3(8) = presc3lab.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc3gen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc3gen.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If presc3gen.Checked = True Then
2:          linha3(7) = "True"
3:      Else
4:          linha3(7) = "False"
5:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub presc4dci_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4dci.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha4(1) = presc4dci.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc4forma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4forma.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha4(2) = presc4forma.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc4dose_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4dose.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha4(3) = presc4dose.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub presc4qty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4qty.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha4(4) = presc4qty.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc4lab_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4lab.SelectedIndexChanged
        On Error GoTo MOSTRARERRO
1:      linha4(8) = presc4lab.Text
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub presc4gen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles presc4gen.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If presc4gen.Checked = True Then
2:          linha4(7) = "True"
3:      Else
4:          linha4(7) = "False"
5:      End If
6:      Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub IndicarComoMostrado(ByVal qualP As Short, ByVal qualA As Short)
        On Error GoTo MOSTRARERRO
        Select Case qualP
            Case 1
                Select Case qualA
                    Case 1
                        a1p1.mostrado = True
                        a1p2.mostrado = True
                        a1p3.mostrado = True
                        a1p4.mostrado = True
                        'a2p1.mostrado = True
                        'a3p1.mostrado = True
                        'a4p1.mostrado = True
                    Case 2
                        a2p1.mostrado = True
                        a2p2.mostrado = True
                        a2p3.mostrado = True
                        a2p4.mostrado = True
                        'a1p1.mostrado = True
                        'a3p1.mostrado = True
                        'a4p1.mostrado = True
                    Case 3
                        a3p1.mostrado = True
                        a3p2.mostrado = True
                        a3p3.mostrado = True
                        a3p4.mostrado = True
                        'a2p1.mostrado = True
                        'a1p1.mostrado = True
                        'a4p1.mostrado = True
                    Case 4
                        a4p1.mostrado = True
                        a4p2.mostrado = True
                        a4p3.mostrado = True
                        a4p4.mostrado = True
                        'a2p1.mostrado = True
                        'a3p1.mostrado = True
                        'a1p1.mostrado = True
                End Select
            Case 2
                Select Case qualA
                    Case 1
                        a1p1.mostrado = True
                        a1p2.mostrado = True
                        a1p3.mostrado = True
                        a1p4.mostrado = True
                        'a2p2.mostrado = True
                        'a3p2.mostrado = True
                        'a4p2.mostrado = True
                    Case 2
                        a2p1.mostrado = True
                        a2p2.mostrado = True
                        a2p3.mostrado = True
                        a2p4.mostrado = True
                        'a1p2.mostrado = True
                        'a3p2.mostrado = True
                        'a4p2.mostrado = True
                    Case 3
                        a3p1.mostrado = True
                        a3p2.mostrado = True
                        a3p3.mostrado = True
                        a3p4.mostrado = True
                        'a2p2.mostrado = True
                        'a1p2.mostrado = True
                        'a4p2.mostrado = True
                    Case 4
                        a4p1.mostrado = True
                        a4p2.mostrado = True
                        a4p3.mostrado = True
                        a4p4.mostrado = True
                        'a2p2.mostrado = True
                        'a3p2.mostrado = True
                        'a1p2.mostrado = True
                End Select
            Case 3
                Select Case qualA
                    Case 1
                        a1p1.mostrado = True
                        a1p2.mostrado = True
                        a1p3.mostrado = True
                        a1p4.mostrado = True
                        'a2p3.mostrado = True
                        'a3p3.mostrado = True
                        'a4p3.mostrado = True
                    Case 2
                        a2p1.mostrado = True
                        a2p2.mostrado = True
                        a2p3.mostrado = True
                        a2p4.mostrado = True
                        'a1p3.mostrado = True
                        'a3p3.mostrado = True
                        'a4p3.mostrado = True
                    Case 3
                        a3p1.mostrado = True
                        a3p2.mostrado = True
                        a3p3.mostrado = True
                        a3p4.mostrado = True
                        'a2p3.mostrado = True
                        'a1p3.mostrado = True
                        'a4p3.mostrado = True
                    Case 4
                        a4p1.mostrado = True
                        a4p2.mostrado = True
                        a4p3.mostrado = True
                        a4p4.mostrado = True
                        'a2p3.mostrado = True
                        'a3p3.mostrado = True
                        'a1p3.mostrado = True
                End Select
            Case 4
                Select Case qualA
                    Case 1
                        a1p1.mostrado = True
                        a1p2.mostrado = True
                        a1p3.mostrado = True
                        a1p4.mostrado = True
                        'a2p4.mostrado = True
                        'a3p4.mostrado = True
                        'a4p4.mostrado = True
                    Case 2
                        a2p1.mostrado = True
                        a2p2.mostrado = True
                        a2p3.mostrado = True
                        a2p4.mostrado = True
                        'a1p4.mostrado = True
                        'a3p4.mostrado = True
                        'a4p4.mostrado = True
                    Case 3
                        a3p1.mostrado = True
                        a3p2.mostrado = True
                        a3p3.mostrado = True
                        a3p4.mostrado = True
                        'a2p4.mostrado = True
                        'a1p4.mostrado = True
                        'a4p4.mostrado = True
                    Case 4
                        a4p1.mostrado = True
                        a4p2.mostrado = True
                        a4p3.mostrado = True
                        a4p4.mostrado = True
                        'a2p4.mostrado = True
                        'a3p4.mostrado = True
                        'a1p4.mostrado = True
                End Select
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub ApagarRepetido(ByVal qualP As Short, ByVal qualA As Short)
        On Error GoTo MOSTRARERRO
        Select Case qualP
            Case 1
                Select Case qualA
                    Case 1
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
                        End If
                    Case 2
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
                        End If
                    Case 3
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

                        End If
                    Case 4
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
                        End If
                End Select
            Case 2
                Select Case qualA
                    Case 1
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
                        End If
                    Case 2
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
                        End If
                    Case 3
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

                        End If
                    Case 4
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
                        End If
                End Select
            Case 3
                Select Case qualA
                    Case 1
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
                        End If
                    Case 2
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
                        End If
                    Case 3
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

                        End If
                    Case 4
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
                        End If
                End Select
            Case 4
                Select Case qualA
                    Case 1
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
                        End If
                    Case 2
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
                        End If
                    Case 3
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

                        End If
                    Case 4
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
                        End If
                End Select
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Dim paramiloidose As Boolean

    Private Sub ParamiloidoseRB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ParamiloidoseRB.CheckedChanged
        On Error GoTo MOSTRARERRO
1:      If ParamiloidoseRB.Checked = True Then
2:          paramiloidose = True
3:      End If
4:      If ParamiloidoseRB.Checked = False Then
6:          paramiloidose = False
7:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub alzheimer(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
        'qual entra no formato aXrow(0) ou aXarray(0)...
1:      MsgBox(qual & " necessita de" & vbCr _
               & "despacho nº. 4250/2007 (ou 25938/08)" & vbCr _
               & "e especialidade de neurologia ou psiquiatria" & vbCr _
               & "para ser comparticipado")
2:      Select Case qual
            Case 1
3:              aviam1.BackColor = Color.Purple
4:          Case 2
5:              aviam2.BackColor = Color.Purple
6:          Case 3
7:              aviam3.BackColor = Color.Purple
8:          Case 4
9:              aviam4.BackColor = Color.Purple
10:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub gastro(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
        'qual entra no formato aXrow(0) ou aXarray(0)...
1:      MsgBox(qual & " necessita de especialidade" & vbCr _
               & "de gastrenterologia, medicina interna, cirurgia geral ou pediatria" & vbCr _
               & "no caso de comparticipado com o despacho nº. 1234/07 (ou 19734/2008 ou 15442/2009)")
2:      Select Case qual
            Case 1
3:              aviam1.BackColor = Color.Purple
4:          Case 2
5:              aviam2.BackColor = Color.Purple
6:          Case 3
7:              aviam3.BackColor = Color.Purple
8:          Case 4
9:              aviam4.BackColor = Color.Purple
10:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub bipolar(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
        'qual entra no formato aXrow(0) ou aXarray(0)...
1:      MsgBox(qual & " necessita de especialidade" & vbCr _
               & "de psiquiatria ou de neurologia" & vbCr _
               & "no caso de comparticipado com o despacho nº. 21094/99")
2:      Select Case qual
            Case 1
3:              aviam1.BackColor = Color.Purple
4:          Case 2
5:              aviam2.BackColor = Color.Purple
6:          Case 3
7:              aviam3.BackColor = Color.Purple
8:          Case 4
9:              aviam4.BackColor = Color.Purple
10:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub espondilite(ByVal qual As Object)
        On Error GoTo MOSTRARERRO
        'qual entra no formato aXrow(0) ou aXarray(0)...
1:      MsgBox(qual & " necessita de especialidade" & vbCr _
             & "de reumatologia ou de medicina interna" & vbCr _
             & "no caso de comparticipado com o despacho nº. 21249/06 ou 14123/2009")
4:      Select Case qual
            Case 1
6:              aviam1.BackColor = Color.Purple
7:          Case 2
8:              aviam2.BackColor = Color.Purple
9:          Case 3
10:             aviam3.BackColor = Color.Purple
11:         Case 4
12:             aviam4.BackColor = Color.Purple
13:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub despacho(ByVal qual As Object, ByVal desp As String)
        On Error GoTo MOSTRARERRO
1:      MsgBox(qual & " pode levar o despacho nº." & desp)
2:      Select Case qual
            Case 1
4:              aviam1.BackColor = Color.Beige
5:          Case 2
6:              aviam2.BackColor = Color.Beige
7:          Case 3
8:              aviam3.BackColor = Color.Beige
9:          Case 4
10:             aviam4.BackColor = Color.Beige
11:             End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub





    Sub AvisarDespachos()
        On Error GoTo MOSTRARERRO
        If P = 1 Then
            If A >= 1 Then
                If Not IsNothing(a1row) Then
                    'If a1row(10) = True Or a1row(11) = True Or a1row(12) = True Or a1row(13) = True Or a1row(14) = True Or a1row(15) = True Then
                    If a1row(9) = True Then
                        If Not a1_4250 = True And av1.nivel <= 3 Then
                            alzheimer(a1row(0))
                            a1_4250 = True
                        End If
                    End If
                    If a1row(10) = True Then
                        If Not a1_1234 = True And av1.nivel <= 3 Then
                            gastro(a1row(0))
                            a1_1234 = True
                        End If
                    End If
                    If a1row(14) = True Then
                        If Not a1_14123 = True And av1.nivel <= 3 Then
                            espondilite(a1row(0))
                            a1_14123 = True
                        End If
                    End If
                    If a1row(15) = True Then
                        If Not a1_1474ad = True And av1.nivel <= 3 Then
                            a1_1474ad = True
                        End If
                    End If
                    If a1row(20) = True Then
                        If Not a1_1474nl = True And av1.nivel <= 3 Then
                            a1_1474nl = True
                        End If
                    End If
                    If a1row(11) = True Then
                        If Not a1_10279 = True And av1.nivel <= 3 Then
                            despacho(a1row(0), "10279/2008, de 11/03")
                            a1_10279 = True
                        End If
                    End If
                    If a1row(12) = True Then
                        If Not a1_10279 = True And av1.nivel <= 3 Then
                            despacho(a1row(0), "10279/2008, de 11/03")
                            a1_10279 = True
                        End If
                    End If
                    If a1row(13) = True Then
                        If Not a1_10910 = True And av1.nivel <= 3 Then
                            despacho(a1row(0), "10910/2009, de 22/04")
                            a1_10910 = True
                        End If
                    End If
                    If a1row(19) = True Then
                        If Not a1_21094 = True And av1.nivel <= 3 Then
                            bipolar(a1row(0))
                            a1_21094 = True
                        End If
                    End If
                End If
            End If
            If A >= 2 And Not IsNothing(a2row) Then
                'If a2row(10) = True Or a2row(11) = True Or a2row(12) = True Or a2row(13) = True Or a2row(14) = True Or a2row(15) = True Then
                If a2row(9) = True Then
                    If Not a2_4250 = True And av2.nivel <= 3 Then
                        alzheimer(a2row(0))
                        a2_4250 = True
                    End If
                End If
                If a2row(10) = True Then
                    If Not a2_1234 = True And av2.nivel <= 3 Then
                        gastro(a2row(0))
                        a2_1234 = True
                    End If
                End If
                If a2row(14) = True Then
                    If Not a2_14123 = True And av2.nivel <= 3 Then
                        espondilite(a2row(0))
                        a2_14123 = True
                    End If
                End If
                If a2row(15) = True Then
                    If Not a2_1474ad = True And av2.nivel <= 3 Then
                        a2_1474ad = True
                    End If
                End If
                If a2row(20) = True Then
                    If Not a2_1474nl = True And av2.nivel <= 3 Then
                        a2_1474nl = True
                    End If
                End If
                If a2row(11) = True Then
                    If Not a2_10279 = True And av2.nivel <= 3 Then
                        despacho(a2row(0), "10279/2008, de 11/03")
                        a2_10279 = True
                    End If
                End If
                If a2row(12) = True Then
                    If Not a2_10279 = True And av2.nivel <= 3 Then
                        despacho(a2row(0), "10279/2008, de 11/03")
                        a2_10279 = True
                    End If
                End If
                If a2row(13) = True Then
                    If Not a2_10910 = True And av2.nivel <= 3 Then
                        despacho(a2row(0), "10910/2009, de 22/04")
                        a2_10910 = True
                    End If
                End If
                If a2row(19) = True Then
                    If Not a2_21094 = True And av2.nivel <= 3 Then
                        bipolar(a2row(0))
                        a2_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 3 And Not IsNothing(a3row) Then
                'If a3row(10) = True Or a3row(11) = True Or a3row(12) = True Or a3row(13) = True Or a3row(14) = True Or a3row(15) = True Then
                If a3row(9) = True Then
                    If Not a3_4250 = True And av3.nivel <= 3 Then
                        alzheimer(a3row(0))
                        a3_4250 = True
                    End If
                End If
                If a3row(10) = True And av3.nivel <= 3 Then
                    If Not a3_1234 = True Then
                        gastro(a3row(0))
                        a3_1234 = True
                    End If
                End If
                If a3row(14) = True And av3.nivel <= 3 Then
                    If Not a3_14123 = True Then
                        espondilite(a3row(0))
                        a3_14123 = True
                    End If
                End If
                If a3row(15) = True Then
                    If Not a3_1474ad = True And av3.nivel <= 3 Then
                        a3_1474ad = True
                    End If
                End If
                If a3row(20) = True Then
                    If Not a3_1474nl = True And av3.nivel <= 3 Then
                        a3_1474nl = True
                    End If
                End If
                If a3row(11) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        despacho(a3row(0), "10279/2008, de 11/03")
                        a3_10279 = True
                    End If
                End If
                If a3row(12) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        despacho(a3row(0), "10279/2008, de 11/03")
                        a3_10279 = True
                    End If
                End If
                If a3row(13) = True And av3.nivel <= 3 Then
                    If Not a3_10910 = True Then
                        despacho(a3row(0), "10910/2009, de 22/04")
                        a3_10910 = True
                    End If
                End If
                If a3row(19) = True Then
                    If Not a3_21094 = True And av3.nivel <= 3 Then
                        bipolar(a3row(0))
                        a3_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 4 And Not IsNothing(a4row) Then
                'If a4row(10) = True Or a4row(11) = True Or a4row(12) = True Or a4row(13) = True Or a4row(14) = True Or a4row(15) = True Then
                If a4row(9) = True Then
                    If Not a4_4250 = True And av4.nivel <= 3 Then
                        alzheimer(a4row(0))
                        a4_4250 = True
                    End If
                End If
                If a4row(10) = True And av4.nivel <= 3 Then
                    If Not a4_1234 = True Then
                        gastro(a4row(0))
                        a4_1234 = True
                    End If
                End If
                If a4row(14) = True And av4.nivel <= 3 Then
                    If Not a4_14123 = True Then
                        espondilite(a4row(0))
                        a4_14123 = True
                    End If
                End If
                If a4row(15) = True Then
                    If Not a4_1474ad = True And av4.nivel <= 3 Then
                        a4_1474ad = True
                    End If
                End If
                If a4row(20) = True Then
                    If Not a4_1474nl = True And av4.nivel <= 3 Then
                        a4_1474nl = True
                    End If
                End If
                If a4row(11) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        despacho(a4row(0), "10279/2008, de 11/03")
                        a4_10279 = True
                    End If
                End If
                If a4row(12) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        despacho(a4row(0), "10279/2008, de 11/03")
                        a4_10279 = True
                    End If
                End If
                If a4row(13) = True And av4.nivel <= 3 Then
                    If Not a4_10910 = True Then
                        despacho(a4row(0), "10910/2009, de 22/04")
                        a4_10910 = True
                    End If
                End If
                If a4row(19) = True Then
                    If Not a4_21094 = True And av4.nivel <= 3 Then
                        bipolar(a4row(0))
                        a4_21094 = True
                    End If
                End If
            End If
        Else
            If A >= 1 Then
                If Not IsNothing(a1row) Then
                    'If a1row(10) = True Or a1row(11) = True Or a1row(12) = True Or a1row(13) = True Or a1row(14) = True Or a1row(15) = True Then
                    If a1row(9) = True Then
                        If Not a1_4250 = True And av1.nivel <= 3 Then
                            alzheimer(a1row(0))
                            a1_4250 = True
                        End If
                    End If
                    If a1row(10) = True And av1.nivel <= 3 Then
                        If Not a1_1234 = True Then
                            gastro(a1row(0))
                            a1_1234 = True
                        End If
                    End If
                    If a1row(14) = True And av1.nivel <= 3 Then
                        If Not a1_14123 = True Then
                            espondilite(a1row(0))
                            a1_14123 = True
                        End If
                    End If
                    If a1row(11) = True And av1.nivel <= 3 Then
                        If Not a1_10279 = True Then
                            despacho(a1row(0), "10279/2008, de 11/03")
                            a1_10279 = True
                        End If
                    End If
                    If a1row(12) = True And av1.nivel <= 3 Then
                        If Not a1_10279 = True Then
                            despacho(a1row(0), "10279/2008, de 11/03")
                            a1_10279 = True
                        End If
                    End If
                    If a1row(13) = True And av1.nivel <= 3 Then
                        If Not a1_10910 = True Then
                            despacho(a1row(0), "10910/2009, de 22/04")
                            a1_10910 = True
                        End If
                    End If
                    If a1row(15) = True And av1.nivel <= 3 Then
                        If Not a1_1474ad = True Then
                            portariado(a1row(0), "1474/2004, de 21/12")
                            a1_1474ad = True
                        End If
                    End If
                    If a1row(20) = True And av1.nivel <= 3 Then
                        If Not a1_1474nl = True Then
                            portariado(a1row(0), "1474/2004, de 21/12")
                            a1_1474nl = True
                        End If
                    End If
                    If a1row(19) = True And av1.nivel <= 3 Then
                        If Not a1_21094 = True Then
                            bipolar(a1row(0))
                            a1_21094 = True
                        End If
                    End If
                End If
            End If
            If A >= 2 And Not IsNothing(a2row) Then
                'If a2row(10) = True Or a2row(11) = True Or a2row(12) = True Or a2row(13) = True Or a2row(14) = True Or a2row(15) = True Then
                If a2row(9) = True And av2.nivel <= 3 Then
                    If Not a2_4250 = True Then
                        alzheimer(a2row(0))
                        a2_4250 = True
                    End If
                End If
                If a2row(10) = True And av2.nivel <= 3 Then
                    If Not a2_1234 = True Then
                        gastro(a2row(0))
                        a2_1234 = True
                    End If
                End If
                If a2row(14) = True And av2.nivel <= 3 Then
                    If Not a2_14123 = True Then
                        espondilite(a2row(0))
                        a2_14123 = True
                    End If
                End If
                If a2row(11) = True And av2.nivel <= 3 Then
                    If Not a2_10279 = True Then
                        despacho(a2row(0), "10279/2008, de 11/03")
                        a2_10279 = True
                    End If
                End If
                If a2row(12) = True And av2.nivel <= 3 Then
                    If Not a2_10279 = True Then
                        despacho(a2row(0), "10279/2008, de 11/03")
                        a2_10279 = True
                    End If
                End If
                If a2row(13) = True And av2.nivel <= 3 Then
                    If Not a2_10910 = True Then
                        despacho(a2row(0), "10910/2009, de 22/04")
                        a2_10910 = True
                    End If
                End If
                If a2row(15) = True And av2.nivel <= 3 Then
                    If Not a2_1474ad = True Then
                        portariado(a2row(0), "1474/2004, de 21/12")
                        a2_1474ad = True
                    End If
                End If
                If a2row(20) = True And av2.nivel <= 3 Then
                    If Not a2_1474nl = True Then
                        portariado(a2row(0), "1474/2004, de 21/12")
                        a2_1474nl = True
                    End If
                End If
                If a2row(19) = True And av2.nivel <= 3 Then
                    If Not a2_21094 = True Then
                        bipolar(a2row(0))
                        a2_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 3 And Not IsNothing(a3row) Then
                'If a3row(10) = True Or a3row(11) = True Or a3row(12) = True Or a3row(13) = True Or a3row(14) = True Or a3row(15) = True Then
                If a3row(9) = True And av3.nivel <= 3 Then
                    If Not a3_4250 = True Then
                        alzheimer(a3row(0))
                        a3_4250 = True
                    End If
                End If
                If a3row(10) = True And av3.nivel <= 3 Then
                    If Not a3_1234 = True Then
                        gastro(a3row(0))
                        a3_1234 = True
                    End If
                End If
                If a3row(14) = True And av3.nivel <= 3 Then
                    If Not a3_14123 = True Then
                        espondilite(a3row(0))
                        a3_14123 = True
                    End If
                End If
                If a3row(11) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        despacho(a3row(0), "10279/2008, de 11/03")
                        a3_10279 = True
                    End If
                End If
                If a3row(12) = True And av3.nivel <= 3 Then
                    If Not a3_10279 = True Then
                        despacho(a3row(0), "10279/2008, de 11/03")
                        a3_10279 = True
                    End If
                End If
                If a3row(13) = True And av3.nivel <= 3 Then
                    If Not a3_10910 = True Then
                        despacho(a3row(0), "10910/2009, de 22/04")
                        a3_10910 = True
                    End If
                End If
                If a3row(15) = True And av3.nivel <= 3 Then
                    If Not a3_1474ad = True Then
                        portariado(a3row(0), "1474/2004, de 21/12")
                        a3_1474ad = True
                    End If
                End If
                If a3row(20) = True And av3.nivel <= 3 Then
                    If Not a3_1474nl = True Then
                        portariado(a3row(0), "1474/2004, de 21/12")
                        a3_1474nl = True
                    End If
                End If
                If a3row(19) = True And av3.nivel <= 3 Then
                    If Not a3_21094 = True Then
                        bipolar(a3row(0))
                        a3_21094 = True
                    End If
                End If
                'End If
            End If
            If A >= 4 And Not IsNothing(a4row) Then
                'If a4row(10) = True Or a4row(11) = True Or a4row(12) = True Or a4row(13) = True Or a4row(14) = True Or a4row(15) = True Then
                If a4row(9) = True And av4.nivel <= 3 Then
                    If Not a4_4250 = True Then
                        alzheimer(a4row(0))
                        a4_4250 = True
                    End If
                End If
                If a4row(10) = True And av4.nivel <= 3 Then
                    If Not a4_1234 = True Then
                        gastro(a4row(0))
                        a4_1234 = True
                    End If
                End If
                If a4row(14) = True And av4.nivel <= 3 Then
                    If Not a4_14123 = True Then
                        espondilite(a4row(0))
                        a4_14123 = True
                    End If
                End If
                If a4row(11) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        despacho(a4row(0), "10279/2008, de 11/03")
                        a4_10279 = True
                    End If
                End If
                If a4row(12) = True And av4.nivel <= 3 Then
                    If Not a4_10279 = True Then
                        despacho(a4row(0), "10279/2008, de 11/03")
                        a4_10279 = True
                    End If
                End If
                If a4row(13) = True And av4.nivel <= 3 Then
                    If Not a4_10910 = True Then
                        despacho(a4row(0), "10910/2009, de 22/04")
                        a4_10910 = True
                    End If
                End If
                If a4row(15) = True And av4.nivel <= 3 Then
                    If Not a4_1474ad = True Then
                        portariado(a4row(0), "1474/2004, de 21/12")
                        a4_1474ad = True
                    End If
                End If
                If a4row(20) = True And av4.nivel <= 3 Then
                    If Not a4_1474nl = True Then
                        portariado(a4row(0), "1474/2004, de 21/12")
                        a4_1474nl = True
                    End If
                End If
                If a4row(19) = True And av4.nivel <= 3 Then
                    If Not a4_21094 = True Then
                        bipolar(a4row(0))
                        a4_21094 = True
                    End If
                End If
                'End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


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


    'a troca de apresentação é avaliada em função da via de administração. cada uma tem muitas formas. aqui se associa a via à forma
    Public Function Via(ByVal Forma As String) As Short
        On Error GoTo MOSTRARERRO
1:      Select Case Forma
            Case "Cápsula"
3:              Via = 10
4:          Case "Cápsula de libertação modificada"
5:              Via = 10
6:          Case "Cápsula de libertação prolongada"
7:              Via = 10
8:          Case "Cápsula gastrorresistente"
9:              Via = 10
10:         Case "Cápsula mole"
11:             Via = 10
12:         Case "Cápsula mole vaginal"
13:             Via = 18
14:         Case "Champô"
15:             Via = 3
16:         Case "Colírio de libertação prolongada"
17:             Via = 9
18:         Case "Colírio, comprimido e solvente para solução"
19:             Via = 9
20:         Case "Colírio, solução"
21:             Via = 9
22:         Case "Colírio, suspensão"
23:             Via = 9
24:         Case "Comprimido"
25:             Via = 10
26:         Case "Comprimido + Suspensão Oral"
27:             Via = 10
28:         Case "Comprimido bucal"
29:             Via = 2
30:         Case "Comprimido bucal mucoadesivo"
31:             Via = 2
32:         Case "Comprimido de libertação modificada"
33:             Via = 10
34:         Case "Comprimido de libertação prolongada"
35:             Via = 10
36:         Case "Comprimido de libertação prolongada revestido por película"
37:             Via = 10
38:         Case "Comprimido dispersível"
39:             Via = 10
40:         Case "Comprimido dispersível ou para mastigar"
41:             Via = 10
42:         Case "Comprimido efervescente"
43:             Via = 10
44:         Case "Comprimido gastrorresistente"
45:             Via = 10
46:         Case "Comprimido orodispersível"
47:             Via = 10
48:         Case "Comprimido para chupar"
49:             Via = 10
50:         Case "Comprimido para mastigar"
51:             Via = 10
52:         Case "Comprimido para suspensão rectal"
53:             Via = 13
54:         Case "Comprimido revestido"
55:             Via = 10
56:         Case "Comprimido revestido por película"
57:             Via = 10
58:         Case "Comprimido solúvel"
59:             Via = 10
60:         Case "Comprimido sublingual"
61:             Via = 15
62:         Case "Comprimido vaginal"
63:             Via = 18
64:         Case "Concentrado e solvente para solução para perfusão"
65:             Via = 12
66:         Case "Concentrado para solução injectável"
67:             Via = 12
68:         Case "Concentrado para solução para perfusão"
69:             Via = 12
70:         Case "Creme"
71:             Via = 3
72:         Case "Creme rectal"
73:             Via = 3
74:         Case "Creme vaginal"
75:             Via = 18
76:         Case "Creme vaginal + Óvulo"
77:             Via = 18
78:         Case "Dispositivo de libertação intra-uterino"
79:             Via = 7
80:         Case "Emplastro medicamentoso"
81:             Via = 3
82:         Case "Emulsão cutânea"
83:             Via = 3
84:         Case "Emulsão e suspensão para emulsão injectável"
85:             Via = 12
86:         Case "Espuma cutânea"
87:             Via = 3
88:         Case "Espuma rectal"
89:             Via = 13
90:         Case "Espuma vaginal"
91:             Via = 18
92:         Case "Gel"
93:             Via = 3
94:         Case "Gel gengival"
95:             Via = 4
96:         Case "Gel nasal"
97:             Via = 8
98:         Case "Gel oftálmico"
99:             Via = 9
100:        Case "Gel oral"
101:            Via = 10
102:        Case "Gel vaginal"
103:            Via = 18
104:        Case "Gotas auriculares, solução"
105:            Via = 1
106:        Case "Gotas auriculares, suspensão"
107:            Via = 1
108:        Case "Gotas nasais, solução"
109:            Via = 8
110:        Case "Gotas orais, solução"
111:            Via = 10
112:        Case "Gotas orais, suspensão"
113:            Via = 10
114:        Case "Granulado"
115:            Via = 10
116:        Case "Granulado de libertação modificada"
117:            Via = 10
118:        Case "Granulado de libertação prolongada"
119:            Via = 10
120:        Case "Granulado de libertação prolongada para suspensão oral"
121:            Via = 10
122:        Case "Granulado efervescente"
123:            Via = 10
124:        Case "Granulado gastrorresistente"
125:            Via = 10
126:        Case "Granulado gastrorresistente para suspensão oral"
127:            Via = 10
128:        Case "Granulado para solução oral"
129:            Via = 10
130:        Case "Granulado para solução oral ou rectal"
131:            Via = 0
132:        Case "Granulado para suspensão oral"
133:            Via = 10
134:        Case "Implante"
135:            Via = 6
136:        Case "Implante em cadeia"
137:            Via = 6
138:        Case "Inserto oftálmico"
139:            Via = 9
140:        Case "Lápis uretral"
141:            Via = 17
142:        Case "Liofilizado oral"
143:            Via = 15
144:        Case "Líquido cutâneo"
145:            Via = 3
146:        Case "Óvulo"
147:            Via = 18
148:        Case "Pasta cutânea"
149:            Via = 3
150:        Case "Pastilha"
151:            Via = 10
152:        Case "Penso impregnado"
153:            Via = 3
154:        Case "Pó cutâneo"
155:            Via = 3
156:        Case "Pó e solvente para para perfusão"
157:            Via = 12
158:        Case "Pó e solvente para solução injectável"
159:            Via = 12
160:        Case "Pó e solvente para solução para perfusão"
161:            Via = 12
162:        Case "Pó e veículo para suspensão injectável"
163:            Via = 12
164:        Case "Pó e veículo para suspensão injectável de libertação prolongada"
165:            Via = 12
166:        Case "Pó e veículo para suspensão oral"
167:            Via = 10
168:        Case "Pó efervescente"
169:            Via = 10
170:        Case "Pó nasal"
171:            Via = 8
172:        Case "Pó oral"
173:            Via = 10
174:        Case "Pó para inalação"
175:            Via = 5
176:        Case "Pó para inalação em recipiente unidose"
177:            Via = 5
178:        Case "Pó para inalação, cápsula"
179:            Via = 5
180:        Case "Pó para pulverização cutânea"
181:            Via = 3
182:        Case "Pó para solução injectável"
183:            Via = 12
184:        Case "Pó para solução injectável ou para perfusão"
185:            Via = 12
186:        Case "Pó para solução oral"
187:            Via = 10
188:        Case "Pó para solução ou para suspensão injectável"
189:            Via = 12
190:        Case "Pó para solução para perfusão"
191:            Via = 12
192:        Case "Pó para solução vaginal"
193:            Via = 18
194:        Case "Pó para suspensão oral"
195:            Via = 10
196:        Case "Pó periodontal"
197:            Via = 14
198:        Case "Pomada"
199:            Via = 3
200:        Case "Pomada oftálmica"
201:            Via = 9
202:        Case "Pomada rectal"
203:            Via = 13
204:        Case "Pomada Rectal + Supositório"
205:            Via = 13
206:        Case "Pomada vaginal"
207:            Via = 18
208:        Case "Sistema de libertação vaginal"
209:            Via = 18
210:        Case "Sistema transdérmico"
211:            Via = 16
212:        Case "Sistema transdérmico "
213:            Via = 16
214:        Case "Solução bucal"
215:            Via = 2
216:        Case "Solução cutânea"
217:            Via = 3
218:        Case "Solução injectável"
219:            Via = 12
220:        Case "Solução injectável ou para perfusão"
221:            Via = 12
222:        Case "Solução oral"
223:            Via = 10
224:        Case "Solução para gargarejar"
225:            Via = 2
226:        Case "Solução para inalação por nebulização"
227:            Via = 5
228:        Case "Solução para inalação por vaporização"
229:            Via = 5
230:        Case "Solução para lavagem da boca"
231:            Via = 2
232:        Case "Solução para lavagem oftálmica"
233:            Via = 9
234:        Case "Solução para perfusão"
235:            Via = 12
236:        Case "Solução para pulverização bucal"
237:            Via = 2
238:        Case "Solução para pulverização cutânea"
239:            Via = 3
240:        Case "Solução para pulverização nasal"
241:            Via = 8
242:        Case "Solução pressurizada para inalação"
243:            Via = 5
244:        Case "Solução rectal"
245:            Via = 13
246:        Case "Solução vaginal"
247:            Via = 18
248:        Case "Solvente/Veículo para uso parentérico"
249:            Via = 12
250:        Case "Supositório"
251:            Via = 13
252:        Case "Suspensão injectável"
253:            Via = 12
254:        Case "Suspensão oral"
255:            Via = 10
256:        Case "Suspensão para inalação por nebulização"
257:            Via = 5
258:        Case "Suspensão para pulverização nasal"
259:            Via = 8
260:        Case "Suspensão pressurizada para inalação"
261:            Via = 5
262:        Case "Suspensão rectal"
263:            Via = 13
264:        Case "Verniz para as unhas medicamentoso"
265:            Via = 3
266:        Case "Xarope"
267:            Via = 10
268:        Case Else
269:            MsgBox("forma farmacêutica desconhecida")
270:            End Select
271:    If Via = 0 Then
272:        MsgBox("granulado para solução oral ou rectal")
273:    End If
        Exit Function
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    'Private Sub butForm3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butForm3.Click
    '       On Error GoTo MOSTRARERRO
    '1:      Form3.BringToFront()
    '        Exit Sub
    'MOSTRARERRO:
    '        Resume Next
    '   End Sub

    Private Sub butEC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEC.Click
        On Error GoTo MOSTRARERRO
1:      EC.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub butEP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEP.Click
        On Error GoTo MOSTRARERRO
1:      EP.Show()
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub portariado(ByVal qual As Object, ByVal port As String)
        On Error GoTo MOSTRARERRO
1:      MsgBox(qual & " pode levar o portaria nº." & port)
2:      Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub codEC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles codEC.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caracta44 As String
2:      Caracta44 = codEC.Text
3:      If Len(codEC.Text) = 7 Then
4:          If Caracta44 Like "#######" Then
5:              limpar4()
6:              codigo4.codigo = codEC.Text
7:              incorporar()
8:              mostrar()
9:              Me.codEC.Focus()
10:             codEC.SelectionStart = 0
11:             codEC.SelectionLength = Len(codEC.Text)
12:         End If
13:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
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
8:      mostradoseqty = ""
9:      mostracomp = ""
10:     mostracompgenports = ""
11:     labelmednome.Text = ""
12:     labelmeddci.Text = ""
13:     labelmedforma.Text = ""
14:     labelmeddoseqty.Text = ""
15:     labelmedcompgenports.Text = ""
16:     labelmednome.BackColor = Color.Transparent
17:     labelmednome.ForeColor = Color.Black
18:     labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Regular)
19:     labelmedcompgenports.Text = ""
20:     labelmedcompgenports.BackColor = Color.Transparent
21:     labelmedcompgenports.ForeColor = Color.Black
22:     labelmedcompgenports.Font = New Font(Me.labelmedcompgenports.Font, FontStyle.Regular)
23:     labelmeddoseqty.Text = ""
24:     labelmeddoseqty.BackColor = Color.Transparent
25:     labelmeddoseqty.ForeColor = Color.Black
26:     labelmeddoseqty.Font = New Font(Me.labelmeddoseqty.Font, FontStyle.Regular)
27:     labelmeddci.Text = ""
28:     labelmeddci.BackColor = Color.Transparent
29:     labelmeddci.ForeColor = Color.Black
30:     labelmeddci.Font = New Font(Me.labelmeddci.Font, FontStyle.Regular)
31:     labelmedforma.Text = ""
32:     labelmedforma.BackColor = Color.Transparent
33:     labelmedforma.ForeColor = Color.Black
34:     labelmedforma.Font = New Font(Me.labelmedforma.Font, FontStyle.Regular)
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
21:                     portec4a = "portaria 1474/2004 (ad)"
22:                     portec4b = ""
23:                 Else
24:                     If codigorow(20) = True Then
25:                         portec4a = "portaria 1474/2004 (nl)"
26:                         portec4b = ""
27:                     Else
28:                         If codigorow(19) = True Then
29:                             portec4a = "despacho 21094/1999"
30:                             portec4b = ""
31:                         Else
32:                             If codigorow(13) = True Then
33:                                 portec4a = "despacho 10910/2009"
34:                                 portec4b = ""
35:                             Else
36:                                 If codigorow(10) = True Then
37:                                     portec4a = "despacho 1234/2007"
38:                                     If codigorow(14) = True Then
39:                                         portec4b = "despacho 14123/2009"
40:                                     Else
41:                                         portec4b = ""
42:                                     End If
43:                                 End If
44:                             End If
45:                         End If
46:                     End If
47:                 End If
48:             End If
49:         End If
50:
51:         labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
52:         labelmeddci.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
53:         labelmedforma.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
54:         labelmeddoseqty.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
55:         labelmedcompgenports.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
56:
57:         mostranome = (codigorow(16))
58:         mostradci = (codigorow(1))
59:         mostraforma = (codigorow(2))
60:         mostradoseqty = "(" & (codigorow(3)) & ") (" & (codigorow(4)) & ")"
61:         mostracompgenports = (codigorow(5)) & "% (" & (genec4) & ") (" & (portec4a) & ") (" & (portec4b) & ")"
62:         If novado = "false" Then
63:         End If
64:     ElseIf mostrado9 = False Then
65:         aviadoNexist(9)
66:         mostrado9 = "True"
67:     End If
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
4:          labelmednome.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
5:          labelmeddci.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
6:          labelmedforma.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
7:          labelmeddoseqty.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
8:          labelmedcompgenports.Font = New Font(Me.labelmednome.Font, FontStyle.Bold)
9:      ElseIf mostrado9 = False Then
10:         aviadoNexist(9)
11:         mostrado9 = "True"
12:     End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub mostrar()
        On Error GoTo MOSTRARERRO
1:      If mostrado9 = False Then
2:          labelmednome.Text = mostranome
3:          labelmeddci.Text = mostradci
4:          labelmedforma.Text = mostraforma
5:          labelmeddoseqty.Text = mostradoseqty
6:          labelmedcompgenports.Text = mostracompgenports
7:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub limparEC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles limparEC.Click
        On Error GoTo MOSTRARERRO
1:      limpar4()
2:      codEC.Text = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub indicar(ByVal which As Short)
1:      On Error GoTo MOSTRARERRO
2:      If Not IsNothing(codigorow) Then
3:          Select Case which
                Case 1
4:                  If av1.mostrado = "true" Then
5:                      If a1row(7) = "true" Then
6:                          gen1.Text = "genérico"
7:                          gen = True
8:                      Else
9:                          gen1.Text = "marca"
10:                         gen = False
11:                     End If
12:                     portaria()
13:                     port1.Text = portimedio
14:                     comp = (a1row(5) * 0.01)
15:                     portcomp1 = portcomp
16:                     intermedio = Replace(a1row(17), ".", ",")
17:                     pvp1.Text = intermedio
18:                     pr1 = Replace(a1row(18), ".", ",")
19:                     If organismo = 48 Or organismo = 49 Then
20:                         pr1 = 1.2 * pr1
21:                     End If
22:                     pr = pr1
23:                     If pr > 0 Then
24:                         intermedio = System.Math.Min(intermedio, pr)
25:                     End If
26:                     comp1.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
27:                 End If
28:             Case 2
29:                 If av2.mostrado = "true" Then
30:                     If a2row(7) = "true" Then
31:                         gen2.Text = "genérico"
32:                         gen = True
33:                     Else
34:                         gen2.Text = "marca"
35:                         gen = False
36:                     End If
37:                     portaria()
38:                     port2.Text = portimedio
39:                     comp = (a2row(5) * 0.01)
40:                     portcomp2 = portcomp
41:                     intermedio = Replace(a2row(17), ".", ",")
42:                     pvp2.Text = intermedio
43:                     pr2 = Replace(a2row(18), ".", ",")
44:                     If organismo = 48 Or organismo = 49 Then
45:                         pr2 = 1.2 * pr2
46:                     End If
47:                     pr = pr2
48:                     If pr > 0 Then
49:                         intermedio = System.Math.Min(intermedio, pr)
50:                     End If
51:                     comp2.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
52:                 End If
53:             Case 3
54:                 If av3.mostrado = "true" Then
55:                     If a3row(7) = "true" Then
56:                         gen3.Text = "genérico"
57:                         gen = True
58:                     Else
59:                         gen3.Text = "marca"
60:                         gen = False
61:                     End If
62:                     portaria()
63:                     port3.Text = portimedio
64:                     comp = (a3row(5) * 0.01)
65:                     portcomp3 = portcomp
66:                     intermedio = Replace(a3row(17), ".", ",")
67:                     pvp3.Text = intermedio
68:                     pr3 = Replace(a3row(18), ".", ",")
69:                     If organismo = 48 Or organismo = 49 Then
70:                         pr3 = 1.2 * pr3
71:                     End If
72:                     pr = pr3
73:                     If pr > 0 Then
74:                         intermedio = System.Math.Min(intermedio, pr)
75:                     End If
76:                     comp3.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
77:                 End If
78:             Case 4
79:                 If av4.mostrado = "true" Then
80:                     If a4row(7) = "true" Then
81:                         gen4.Text = "genérico"
82:                         gen = True
83:                     Else
84:                         gen4.Text = "marca"
85:                         gen = False
86:                     End If
87:                     portaria()
88:                     port4.Text = portimedio
89:                     comp = (a4row(5) * 0.01)
90:                     portcomp4 = portcomp
91:                     intermedio = Replace(a4row(17), ".", ",")
92:                     pvp4.Text = intermedio
93:                     pr4 = Replace(a4row(18), ".", ",")
94:                     If organismo = 48 Or organismo = 49 Then
95:                         pr4 = 1.2 * pr4
96:                     End If
97:                     pr = pr4
98:                     If pr > 0 Then
99:                         intermedio = System.Math.Min(intermedio, pr)
100:                    End If
101:                    comp4.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
102:                End If
103:            Case 5
104:                If av5.mostrado = "true" Then
105:                    If a5row(7) = "true" Then
106:                        gen5.Text = "genérico"
107:                        gen = True
108:                    Else
109:                        gen5.Text = "marca"
110:                        gen = False
111:                    End If
112:                    portaria()
113:                    port5.Text = portimedio
114:                    comp = (a5row(5) * 0.01)
115:                    portcomp5 = portcomp
116:                    intermedio = Replace(a5row(17), ".", ",")
117:                    pvp5.Text = intermedio
118:                    pr5 = Replace(a5row(18), ".", ",")
119:                    If organismo = 48 Or organismo = 49 Then
120:                        pr5 = 1.2 * pr5
121:                    End If
122:                    pr = pr5
123:                    If pr > 0 Then
124:                        intermedio = System.Math.Min(intermedio, pr)
125:                    End If
126:                    comp5.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
127:                End If
128:            Case 6
129:                If av6.mostrado = "true" Then
130:                    If a6row(7) = "true" Then
131:                        gen6.Text = "genérico"
132:                        gen = True
133:                    Else
134:                        gen6.Text = "marca"
135:                        gen = False
136:                    End If
137:                    portaria()
138:                    port6.Text = portimedio
139:                    comp = (a6row(5) * 0.01)
140:                    portcomp6 = portcomp
141:                    intermedio = Replace(a6row(17), ".", ",")
142:                    pvp6.Text = intermedio
143:                    pr6 = Replace(a6row(18), ".", ",")
144:                    If organismo = 48 Or organismo = 49 Then
145:                        pr6 = 1.2 * pr6
146:                    End If
147:                    pr = pr6
148:                    If pr > 0 Then
149:                        intermedio = System.Math.Min(intermedio, pr)
150:                    End If
151:                    comp6.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
152:                End If
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
31:         portimedio = "147469"
32:     End If
33:
34:     If codigorow(19) = True Then
35:         portimedio = "21094"
36:     End If
37:
38:     If codigorow(20) = True Then
39:         portimedio = "1474100"
40:     End If
41:     If codigorow(9) = False And codigorow(10) = False And codigorow(11) = False And codigorow(12) = False And codigorow(13) = False _
     And codigorow(14) = False And codigorow(15) = False And codigorow(19) = False And codigorow(20) = False Then
42:         portimedio = "não"
43:     End If
44:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB portaria: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Function calculo(ByVal org As Short, ByVal gen As Boolean, ByVal comp As Double, ByVal intermedio As Double) As Double
1:      On Error GoTo MOSTRARERRO
2:
3:      Select Case org
            Case 1 'tipo 10
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
23:             calculo = intermedio * 0.85
24:         Case 48 'tipo 15
25:             If comp > 0 Then
26:                 If gen = "true" Then
27:                     calculo = intermedio
28:                 Else
29:                     calculo = (System.Math.Min(1, (comp + 0.15))) * intermedio
30:                 End If
31:             Else
32:                 calculo = 0
33:             End If
34:         Case 45
35:             calculo = intermedio * (System.Math.Max(comp, portcomp))
36:         Case 49
37:             If gen = "true" Then
38:                 calculo = intermedio
39:             Else
40:                 calculo = System.Math.Min(1, (System.Math.Max((portcomp + 0.15), (comp + 0.15)))) * intermedio
41:             End If
42:             End Select
43:
44:     'não tenho nada para o tipo 19 nem para os organismos 02_ADSE,12_BancNt,15_IASFA,17_GNR,18_PSP,25_SAMSq,59_ADSEdipl,R5_CGD
45:
46:     Exit Function
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
23:     If pvp5.Text <> "" Then
24:         pvp5v = Replace(pvp5.Text, ".", ",")
25:         pvp5val = Convert.ToDouble(pvp5v)
26:     Else : pvp5val = 0
27:     End If
28:     If pvp6.Text <> "" Then
29:         pvp6v = Replace(pvp6.Text, ".", ",")
30:         pvp6val = Convert.ToDouble(pvp6v)
31:     Else : pvp6val = 0
32:     End If
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
3:      If comp1.Text <> "" Then
4:          comp1v = Replace(comp1.Text, ".", ",")
5:          comp1val = Convert.ToDouble(comp1v)
6:      Else : comp1val = 0
7:      End If
8:      If comp2.Text <> "" Then
9:          comp2v = Replace(comp2.Text, ".", ",")
10:         comp2val = Convert.ToDouble(comp2v)
11:     Else : comp2val = 0
12:     End If
13:     If comp3.Text <> "" Then
14:         comp3v = Replace(comp3.Text, ".", ",")
15:         comp3val = Convert.ToDouble(comp3v)
16:     Else : comp3val = 0
17:     End If
18:     If comp4.Text <> "" Then
19:         comp4v = Replace(comp4.Text, ".", ",")
20:         comp4val = Convert.ToDouble(comp4v)
21:     Else : comp4val = 0
22:     End If
23:     If comp5.Text <> "" Then
24:         comp5v = Replace(comp5.Text, ".", ",")
25:         comp5val = Convert.ToDouble(comp5v)
26:     Else : comp5val = 0
27:     End If
28:     If comp6.Text <> "" Then
29:         comp6v = Replace(comp6.Text, ".", ",")
30:         comp6val = Convert.ToDouble(comp6v)
31:     Else : comp6val = 0
32:     End If
33:     somaComp = comp1val + comp2val + comp3val + comp4val + comp5val + comp6val
34:     SomarComp = somaComp
35:     Exit Function
MOSTRARERRO:
        MsgBox("SUB somarcomp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Sub somas()
1:      On Error GoTo MOSTRARERRO
2:      agrupar()
3:      SomarPVP()
4:      SomarComp()
5:      totalPVP.Text = SomarPVP()
6:      totalComp.Text = SomarComp()
7:      Exit Sub
MOSTRARERRO:
        MsgBox("SUB somas: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but1474_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(15) = "true" And a1row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          ElseIf a1row(20) = "true" And a1row(5) <> 0 Then
7:              portcomp = 1
8:          Else
9:              portcomp = 0
10:         End If
11:         indicar(1)
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(15) = "true" And a2row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          ElseIf a2row(20) = "true" And a2row(5) <> 0 Then
7:              portcomp = 1
8:          Else
9:              portcomp = 0
10:         End If
11:         indicar(2)
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(15) = "true" And a3row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          ElseIf a3row(20) = "true" And a3row(5) <> 0 Then
7:              portcomp = 1
8:          Else
9:              portcomp = 0
10:         End If
11:         indicar(3)
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_04.Checked Then
3:          If a4row(15) = "true" And a4row(5) <> 0 Then
4:              portcomp = Replace(0.69, ".", ",")
5:          ElseIf a4row(20) = "true" And a4row(5) <> 0 Then
6:              portcomp = 1
7:          Else
8:              portcomp = 0
9:          End If
10:         indicar(4)
11:         somas()
12:     End If
13:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(15) = "true" And a5row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          ElseIf a5row(20) = "true" And a5row(5) <> 0 Then
7:              portcomp = 1
8:          Else
9:              portcomp = 0
10:         End If
11:         indicar(5)
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1474_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(15) = "true" And a6row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          ElseIf a6row(20) = "true" And a6row(5) <> 0 Then
7:              portcomp = 1
8:          Else
9:              portcomp = 0
10:         End If
11:         indicar(6)
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(10) = "true" And a1row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(1)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(10) = "true" And a2row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(2)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(10) = "true" And a3row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(3)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(10) = "true" And a4row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(4)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(10) = "true" And a5row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(5)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but1234_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(10) = "true" And a6row(5) <> 0 Then
5:              portcomp = Replace(0.95, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(6)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(11) = "true" Or a1row(12) = "true" Then
5:              If a1row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(1)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(11) = "true" Or a2row(12) = "true" Then
5:              If a2row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(2)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(11) = "true" Or a3row(12) = "true" Then
5:              If a3row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(3)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(11) = "true" Or a4row(12) = "true" Then
5:              If a4row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(4)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(11) = "true" Or a5row(12) = "true" Then
5:              If a5row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(5)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10279_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(11) = "true" Or a6row(12) = "true" Then
5:              If a6row(5) <> 0 Then
6:                  portcomp = Replace(0.95, ".", ",")
7:              Else
8:                  portcomp = 0
9:              End If
10:             indicar(6)
11:             somas()
12:         End If
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but14123_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(14) = "true" And a1row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(1)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(14) = "true" And a2row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(2)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(14) = "true" And a3row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(3)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(14) = "true" And a4row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(4)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(14) = "true" And a5row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(5)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but14123_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(14) = "true" And a6row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(6)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(13) = "true" And a1row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(1)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(13) = "true" And a2row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(2)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(13) = "true" And a3row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(3)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(13) = "true" And a4row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(4)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(13) = "true" And a5row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(5)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but10910_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(13) = "true" And a6row(5) <> 0 Then
5:              portcomp = Replace(0.69, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(6)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(19) = "true" And a1row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(19) = "true" And a2row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(19) = "true" And a3row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but21094_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(19) = "true" And a4row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(19) = "true" And a5row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but21094_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(19) = "true" And a6row(5) <> 0 Then
5:              portcomp = 1
6:          Else
7:              portcomp = 0
8:          End If
9:          For i = 1 To 6
10:             indicar(i)
11:         Next
12:         somas()
13:     End If
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_01.Checked Then
3:          portimedio = port1.Text
4:          If a1row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(1)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_02.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_02.Checked Then
3:          portimedio = port2.Text
4:          If a2row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(2)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_03.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_03.Checked Then
3:          portimedio = port3.Text
4:          If a3row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(3)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_04.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_04.Checked Then
3:          portimedio = port4.Text
4:          If a4row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(4)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_05.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_05.Checked Then
3:          portimedio = port5.Text
4:          If a5row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(5)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_06.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but4250_06.Checked Then
3:          portimedio = port6.Text
4:          If a6row(9) = "true" Then
5:              portcomp = Replace(0.37, ".", ",")
6:          Else
7:              portcomp = 0
8:          End If
9:          indicar(6)
10:         somas()
11:     End If
12:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but01.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but01.Checked Then
3:          deslabelar()
4:          organismo = 1
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

    Private Sub but48_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but48.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but48.Checked Then
3:          deslabelar()
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

    Private Sub but41_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but41.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but41.Checked Then
3:          deslabelar()
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

    Private Sub but46_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but46.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but46.Checked Then
3:          deslabelar()
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


    Private Sub but42_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but42.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but42.Checked Then
3:          deslabelar()
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


    Private Sub but67_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but67.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but67.Checked Then
3:          deslabelar()
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

    Private Sub butDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDS.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If butDS.Checked Then
3:          deslabelar()
4:          organismo = 23
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


    Private Sub but49_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but49.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but49.Checked Then
3:          deslabelar()
4:          organismo = 49
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "portaria 1474/2004"
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

    Private Sub but45_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but45.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but45.Checked Then
3:          deslabelar()
4:          organismo = 45
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "portaria 1474/2004"
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

    Private Sub but59_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but59.CheckedChanged
1:      On Error GoTo MOSTRARERRO
2:      If but59.Checked Then
3:          deslabelar()
4:          organismo = 59
5:          organismus(organismo)
6:          labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
7:          labelatribuido.Text = "portaria 1474/2004"
8:          aviam1.Focus()
9:      End If
10:     For i = 1 To 6
11:         indicar(i)
12:         somas()
13:     Next
14:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB but59_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub tirarports()
1:      On Error GoTo MOSTRARERRO
2:      but1474_01.Checked = False
3:      but1474_02.Checked = False
4:      but1474_03.Checked = False
5:      but1474_04.Checked = False
6:      but1474_05.Checked = False
7:      but1474_06.Checked = False
8:      but1234_01.Checked = False
9:      but1234_02.Checked = False
10:     but1234_03.Checked = False
11:     but1234_04.Checked = False
12:     but1234_05.Checked = False
13:     but1234_06.Checked = False
14:     but4250_01.Checked = False
15:     but4250_02.Checked = False
16:     but4250_03.Checked = False
17:     but4250_04.Checked = False
18:     but4250_05.Checked = False
19:     but4250_06.Checked = False
20:     but14123_01.Checked = False
21:     but14123_02.Checked = False
22:     but14123_03.Checked = False
23:     but14123_04.Checked = False
24:     but14123_05.Checked = False
25:     but14123_06.Checked = False
26:     but21094_01.Checked = False
27:     but21094_02.Checked = False
28:     but21094_03.Checked = False
29:     but21094_04.Checked = False
30:     but21094_05.Checked = False
31:     but21094_06.Checked = False
32:     but10279_01.Checked = False
33:     but10279_02.Checked = False
34:     but10279_03.Checked = False
35:     but10279_04.Checked = False
36:     but10279_05.Checked = False
37:     but10279_06.Checked = False
38:     but10910_01.Checked = False
39:     but10910_02.Checked = False
40:     but10910_03.Checked = False
41:     but10910_04.Checked = False
42:     but10910_05.Checked = False
43:     but10910_06.Checked = False
44:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB tirarports: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub deslabelar()
1:      On Error GoTo MOSTRARERRO
2:      labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Regular)
3:      labelatribuido.Text = ""
4:      Exit Sub
MOSTRARERRO:
        MsgBox("SUB deslabelar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub organismus(ByVal umdeles As Short)
1:      On Error GoTo MOSTRARERRO
2:      Select Case umdeles
            Case 1, 46, 41, 42, 67, 23
4:              tirarports()
5:          Case 49, 45, 59
6:              but1474_01.Checked = True
7:              but1474_02.Checked = True
8:              but1474_03.Checked = True
9:              but1474_04.Checked = True
10:             but1474_05.Checked = True
11:             but1474_06.Checked = True
12:             End Select
13:     Exit Sub
MOSTRARERRO:
        MsgBox("SUB organismus: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


End Class


