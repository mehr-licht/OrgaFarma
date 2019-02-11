Public Class FormNoCode
    Dim NCcodigo As New meds
    Dim foco As String
    Dim mostradciNC As String
    Dim mostraffNC As String
    Dim mostradoseNC As String
    Dim mostraqtyNC As String
    Dim A As Short
    Dim NCa1array As New ArrayList
    Dim NCa2array As New ArrayList
    Dim NCa3array As New ArrayList
    Dim NCa4array As New ArrayList
    Dim NCa1row As basededadosDataSet.infarmedRow
    Dim NCa2row As basededadosDataSet.infarmedRow
    Dim NCa3row As basededadosDataSet.infarmedRow
    Dim NCa4row As basededadosDataSet.infarmedRow
    Dim DSNC As New basededadosDataSet
    Dim NCAviado1 As New meds
    Dim NCAviado2 As New meds
    Dim NCAviado3 As New meds
    Dim NCAviado4 As New meds
    Dim NCcodigorow As basededadosDataSet.infarmedRow
    Dim NCnovado As Boolean
    Dim NCav1 As New avaliacao
    Dim NCav2 As New avaliacao
    Dim NCav3 As New avaliacao
    Dim NCav4 As New avaliacao

    'exemplo loops com alteraçao nome variavel
    'For b = 1 To Count
    'Dim image As New LinkedResource(Server.MapPath(myArray(b - 1)))
    'image.ContentId = "imageContentId_" + b.ToString
    ''add the LinkedResource to the appropriate view'
    'htmlView.LinkedResources.Add(image)
    'Next b

    'outro exemplo loops com alteraçao nome variavel
    'Dim output As New StringBuilder("")
    'For i as Integer = 0 To rows.Count - 1
    'output.append("Applicant" + i.ToString())
    'Foreach(col as DataColumn in dt.Columns)  ' The datatable where your rows are
    'Dim colName As String = col.ColumnName
    '   output.append(colName & "=" & rows(i)(colName).ToString())
    'Next
    'If i < rows.Count - 1 Then output.Append("|")
    'Next

    Dim infarmedTANC As New basededadosDataSetTableAdapters.infarmedTableAdapter
    Dim rowNC As DataRow
    Dim infarmedTAs As Integer = infarmedTANC.Fill(DSNC.infarmed)

    Private Sub FormNoCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo MOSTRARERRO
        Me.KeyPreview = True
        Me.WindowState = FormWindowState.Maximized
        NClimpar()
        Me.NCaviam1.Focus()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub FormNoCode_Load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub NClimpar()
        On Error GoTo MOSTRARERRO
        NCaviam1.Text = ""
        NCaviam2.Text = ""
        NCaviam3.Text = ""
        NCaviam4.Text = ""
        'NC1dcilabel.Text = ""
        'NC2dcilabel.Text = ""
        'NC3dcilabel.Text = ""
        'NC4dcilabel.Text = ""
        'NC1doselabel.Text = ""
        'NC2doselabel.Text = ""
        'NC3doselabel.Text = ""
        'NC4doselabel.Text = ""
        'NC1fflabel.Text = ""
        'NC2fflabel.Text = ""
        'NC3fflabel.Text = ""
        'NC4fflabel.Text = ""
        'NC1qtylabel.Text = ""
        'NC2qtylabel.Text = ""
        'NC3qtylabel.Text = ""
        'NC4qtylabel.Text = ""
        nca1rt.Clear()
        nca2rt.Clear()
        nca3rt.Clear()
        nca4rt.Clear()
        NC1rt.Clear()
        NC2rt.Clear()
        NC3rt.Clear()
        NC4rt.Clear()
        NCgb1.BackColor = SystemColors.Control
        NCaviam1.BackColor = Color.White
        nca1rt.BackColor = SystemColors.Control
        NC1rt.BackColor = SystemColors.Control
        NCgb2.BackColor = SystemColors.Control
        NCaviam2.BackColor = Color.White
        nca2rt.BackColor = SystemColors.Control
        NC2rt.BackColor = SystemColors.Control
        NCgb3.BackColor = SystemColors.Control
        NCaviam3.BackColor = Color.White
        nca3rt.BackColor = SystemColors.Control
        NC3rt.BackColor = SystemColors.Control
        NCgb4.BackColor = SystemColors.Control
        NCaviam4.BackColor = Color.White
        nca4rt.BackColor = SystemColors.Control
        NC4rt.BackColor = SystemColors.Control
        Me.NCaviam1.Focus()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub NClimpar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub data_Keyspace(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        On Error GoTo MOSTRARERRO
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
            NClimpar()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub data_Keyspace: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub data_KeysEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        On Error GoTo MOSTRARERRO
        If e.KeyCode = Keys.Enter Then
            Select Case foco
                Case "NCaviam1"
55:                 If NCaviam1.Text = "" Then
56:                     Beep()
57:                 Else
58:                     Me.NCaviam2.Focus()
59:                 End If
60:
61:
91:             Case "NCaviam2"
92:                 If NCaviam2.Text = "" Then
93:                     Me.NCaviam2.Text = "0"
94:                     Me.NCaviam3.Text = "0"
95:                     Me.NCaviam4.Text = "0"
a95:                Else
98:                     Me.NCaviam3.Focus()
99:                 End If
100:            Case "NCaviam3"
101:                If NCaviam3.Text = "" Then
102:                    Me.NCaviam3.Text = "0"
103:                    Me.NCaviam4.Text = "0"
a103:
104:                Else
106:                    Me.NCaviam4.Focus()
107:                End If
a107:           Case "NCaviam4"

z92:                If NCaviam4.Text = "" Then

z93:                    Me.NCaviam4.Text = "0"
                        Me.NCaviam1.Focus()
z96:                Else
b110:                   If NCaviam4.Text >= 1111111 And NCaviam4.Text <= 9999999 Then
e110:                       Me.NCaviam1.Focus()
f110:                   Else
g110:                       Beep()
h110:                       Me.NCaviam4.Text = ""
i110:                       Me.NCaviam4.Focus()
z99:                    End If
                    End If
113:                End Select
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub data_KeysEnter: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub focoKeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
1:      foco = Me.ActiveControl.Name()
    End Sub







    'os próximos 4 são os validadores que estão a funcionar. fazem saltar o focus quando se inserem 7 caracteres. e lançam a comparação
    Private Sub NCaviam1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NCaviam1.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractnc1 As String
2:      Caractnc1 = NCaviam1.Text
3:      If Len(NCaviam1.Text) = 7 Then
4:          If Caractnc1 Like "#######" Then
5:              NCAviado1.codigo = NCaviam1.Text
                NCav1.mostrado = "True"
                'A = 1
                NCa1row = DSNC.infarmed.FindBycode(NCAviado1.codigo)
                NCcodigorow = DSNC.infarmed.FindBycode(NCAviado1.codigo)
                NCa1array.Add(NCa1row)
                'NCirbuscar()
                'NCincorporar()
                NCindicar(1)
                'Me.av2.Focus()
6:              'Me.aviam2.Focus()
            Else
                With NC1rt
                    .Text = "código inválido"
                    .ForeColor = Color.Gray
                End With
7:          End If
            '  Else
            '     With NC1rt
            '.Text = "código incompleto"
            '.ForeColor = Color.Gray
            'End With
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub NCaviam1_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub NCaviam2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NCaviam2.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractnc2 As String
2:      Caractnc2 = NCaviam2.Text
3:
4:      If Len(NCaviam2.Text) = 7 Then
            If Caractnc2 Like "#######" Then
5:              NCAviado2.codigo = NCaviam2.Text
                NCav2.mostrado = "True"
                A = 2
                NCa2row = DSNC.infarmed.FindBycode(NCAviado2.codigo)
                NCcodigorow = DSNC.infarmed.FindBycode(NCAviado2.codigo)
                NCa2array.Add(NCa2row)
                'NCirbuscar()
                'NCincorporar()
                NCindicar(2)
                'Me.av2.Focus()
6:              'Me.aviam2.Focus()
7:          Else
                With NC2rt
                    .Text = "código inválido"
                    .ForeColor = Color.Gray
                End With
            End If
            ' Else
            '     With NC2rt
            ' .Text = "código incompleto"
            ' .ForeColor = Color.Gray
            ' End With
8:      End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub NCaviam2_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub NCaviam3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NCaviam3.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractnc3 As String
2:      Caractnc3 = NCaviam3.Text
3:
4:      If Len(NCaviam3.Text) = 7 Then
            If Caractnc3 Like "#######" Then
5:              NCAviado3.codigo = NCaviam3.Text
                NCav3.mostrado = "True"
                A = 3
                NCa3row = DSNC.infarmed.FindBycode(NCAviado3.codigo)
                NCcodigorow = DSNC.infarmed.FindBycode(NCAviado3.codigo)
                NCa3array.Add(NCa3row)
                'NCirbuscar()
                'NCincorporar()
                NCindicar(3)
            Else
                With NC3rt
                    .Text = "código inválido"
                    .ForeColor = Color.Gray
                End With
            End If
            '  Else
            '      With NC3rt
            ' .Text = "código incompleto"
            ' .ForeColor = Color.Gray
            ' End With
8:      End If
        Exit Sub
10:
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub NCaviam3_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub NCaviam4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NCaviam4.TextChanged
        On Error GoTo MOSTRARERRO
1:      Dim Caractnc4 As String
2:      Caractnc4 = NCaviam4.Text
3:
4:      If Len(NCaviam4.Text) = 7 Then
            If Caractnc4 Like "#######" Then
5:              NCAviado4.codigo = NCaviam4.Text
                NCav4.mostrado = "True"
                A = 4
                NCa4row = DSNC.infarmed.FindBycode(NCAviado4.codigo)
                NCcodigorow = DSNC.infarmed.FindBycode(NCAviado4.codigo)
                NCa4array.Add(NCa4row)
                'NCirbuscar()
                'NCincorporar()
                NCindicar(4)
                'Me.av2.Focus()
6:              'Me.aviam2.Focus()
7:          Else
                With NC4rt
                    .Text = "código inválido"
                    .ForeColor = Color.Gray
                End With
            End If
            'Else
            '   With NC4rt
            '.Text = "código incompleto"
            '.ForeColor = Color.Gray
            'End With
8:      End If
        Exit Sub
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub NCaviam4_TextChanged(: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub




    Sub NCindicar(ByVal which As Short)
1:      On Error GoTo MOSTRARERRO
2:      ' If Not IsNothing(NCcodigorow) Then
3:      Select Case which
            Case 1
4:              If NCav1.mostrado = "true" Then
                    If Not IsNothing(NCa1row) Then
                        Dim tamanhoantes As Single
                        tamanhoantes = 0
5:                      'NC1dcilabel.Text = NCa1row(1) & ", " & NCa1row(4) & ", " & NCa1row(3) & ", " & NCa1row(5) & " unidade(s)"
7:
                        If IsNumeric(NCa1row(5)) Then
                            If NCa1row(25) > 0 Then
                                NC1rt.Text = NCa1row(25) & ", " & NCa1row(1) & ", " & NCa1row(4) & ", " & NCa1row(3) & ", " & NCa1row(5) & " unidade(s)"
                            Else
                                NC1rt.Text = "semCNPEM" & ", " & NCa1row(1) & ", " & NCa1row(4) & ", " & NCa1row(3) & ", " & NCa1row(5) & " unidade(s)"
                            End If
                        Else
                            If NCa1row(25) > 0 Then
                                NC1rt.Text = NCa1row(25) & ", " & NCa1row(1) & ", " & NCa1row(4) & ", " & NCa1row(3) & ", " & NCa1row(5)
                            Else
                                NC1rt.Text = "semCNPEM" & ", " & NCa1row(1) & ", " & NCa1row(4) & ", " & NCa1row(3) & ", " & NCa1row(5)
                            End If
                        End If

                        With NC1rt
                            .SelectionStart = 0
                            .SelectionLength = 10
                            .SelectionColor = Color.Black

                            .SelectionStart = 10
                            .SelectionLength = Len(NCa1row(1)) + 2
                            .SelectionColor = Color.Blue

                            .SelectionStart = Len(NCa1row(1)) + 12
                            .SelectionLength = 2 + Len(NCa1row(4))
                            .SelectionColor = Color.Brown

                            .SelectionStart = 14 + Len(NCa1row(4)) + Len(NCa1row(1))
                            .SelectionLength = 2 + Len(NCa1row(3))
                            .SelectionColor = Color.Purple

                            .SelectionStart = 16 + Len(NCa1row(4)) + Len(NCa1row(1)) + Len(NCa1row(3))
                            .SelectionLength = Len(NCa1row(5)) + 13
                            .SelectionColor = Color.DarkGoldenrod
                        End With
                        nca1rt.Text = NCa1row(2) & "; GH0" & NCa1row(7) & "; " & NCa1row(6) & "%; PVP=" & NCa1row(8) & "; PR=" & NCa1row(9) & "; TOP5=" & NCa1row(19)
                        If NCa1row(20) = True Then
                            nca1rt.Text += "; DCI"
                        End If
                        If NCa1row(12) = True Then
                            nca1rt.Text += "; desp.13020/2011"
                            If NCa1row(20) = True Then
                                tamanhoantes = 5
                            End If
                            With nca1rt
                                .SelectionStart = tamanhoantes + 28 + Len(NCa1row(2)) + Len(NCa1row(7).ToString) + Len(NCa1row(6).ToString) + Len(NCa1row(8).ToString) + Len(NCa1row(9).ToString) + Len(NCa1row(19).ToString)
                                .SelectionLength = 15
                                .SelectionColor = Color.Red
                                .SelectionStart = 7 + Len(NCa1row(2)) + Len(NCa1row(7).ToString)
                                .SelectionLength = 2
                                .SelectionColor = Color.Red
                            End With
                        End If
                        If NCa1row(13) = True Then
                            nca1rt.Text += "; desp.1234/2007"
                        End If
                        If NCa1row(15) = True Then
                            nca1rt.Text += "; desp.10279/2008"
                        End If
                        If NCa1row(17) = True Then
                            nca1rt.Text += "; desp.10910/2009"
                        End If
                        If NCa1row(18) = True Then
                            nca1rt.Text += "; desp.14123/2009"
                        End If
                        If NCa1row(14) = True Then
                            nca1rt.Text += "; desp.21094/99"
                        End If
                        If NCa1row(21) = True Then
                            nca1rt.Text += "; lei 6/2010"
                        End If
                        If NCa1row(20) = False Then
                            nca1rt.SelectionStart = 0
                            nca1rt.SelectionLength = Len(NCa1row(2))
                            nca1rt.SelectionColor = Color.Purple
                        End If
                        If NCa1row(8) <> 0 And NCa1row(19) <> 0 Then
                            If NCa1row(8) > NCa1row(19) Then
                                nca1rt.SelectionStart = tamanhoantes + 10 + Len(NCa1row(2)) + Len(NCa1row(7).ToString) + Len(NCa1row(6).ToString)
                                nca1rt.SelectionLength = Len(NCa1row(8).ToString) + 4
                                nca1rt.SelectionColor = Color.Red
                            Else
                                nca1rt.SelectionStart = tamanhoantes + 10 + Len(NCa1row(2)) + Len(NCa1row(7).ToString) + Len(NCa1row(6).ToString)
                                nca1rt.SelectionLength = Len(NCa1row(8).ToString) + 4
                                nca1rt.SelectionColor = Color.Green
                            End If
                        End If
                    Else
                        With NC1rt
                            .Text = "aviado código desconhecido"
                            .ForeColor = Color.Gray
                        End With
                    End If
                End If
11:         Case 2
12:             If NCav2.mostrado = "true" Then
                    If Not IsNothing(NCa2row) Then
                        Dim tamanhoantes As Single
                        tamanhoantes = 0
                        'NC2dcilabel.Text = NCa2row(1) & ", " & NCa2row(4) & ", " & NCa2row(3) & ", " & NCa2row(5) & " unidade(s)"
                        If IsNumeric(NCa2row(5)) Then
                            If NCa2row(25) > 0 Then
                                NC2rt.Text = NCa2row(25) & ", " & NCa2row(1) & ", " & NCa2row(4) & ", " & NCa2row(3) & ", " & NCa2row(5) & " unidade(s)"
                            Else
                                NC2rt.Text = "semCNPEM" & ", " & NCa2row(1) & ", " & NCa2row(4) & ", " & NCa2row(3) & ", " & NCa2row(5) & " unidade(s)"
                            End If
                        Else
                            If NCa2row(25) > 0 Then
                                NC2rt.Text = NCa2row(25) & ", " & NCa2row(1) & ", " & NCa2row(4) & ", " & NCa2row(3) & ", " & NCa2row(5)
                            Else
                                NC2rt.Text = "semCNPEM" & ", " & NCa2row(1) & ", " & NCa2row(4) & ", " & NCa2row(3) & ", " & NCa2row(5)
                            End If
                        End If
                        With NC2rt
                            .SelectionStart = 0
                            .SelectionLength = 10
                            .SelectionColor = Color.Black

                            .SelectionStart = 10
                            .SelectionLength = Len(NCa2row(1)) + 2
                            .SelectionColor = Color.Blue

                            .SelectionStart = Len(NCa2row(1)) + 12
                            .SelectionLength = 2 + Len(NCa2row(4))
                            .SelectionColor = Color.Brown

                            .SelectionStart = 14 + Len(NCa2row(4)) + Len(NCa2row(1))
                            .SelectionLength = 2 + Len(NCa2row(3))
                            .SelectionColor = Color.Purple

                            .SelectionStart = 16 + Len(NCa2row(4)) + Len(NCa2row(1)) + Len(NCa2row(3))
                            .SelectionLength = Len(NCa2row(5)) + 13
                            .SelectionColor = Color.DarkGoldenrod
                        End With
                        nca2rt.Text = NCa2row(2) & "; GH0" & NCa2row(7) & "; " & NCa2row(6) & "%; PVP=" & NCa2row(8) & "; PR=" & NCa2row(9) & "; TOP5=" & NCa2row(19)
                        If NCa2row(20) = True Then
                            nca2rt.Text += "; DCI"
                        End If
                        If NCa2row(12) = True Then
                            nca2rt.Text += "; desp.13020/2011"
                            If NCa2row(20) = True Then
                                tamanhoantes = 5
                            End If
                            With nca2rt
                                .SelectionStart = tamanhoantes + 28 + Len(NCa2row(2)) + Len(NCa2row(7).ToString) + Len(NCa2row(6).ToString) + Len(NCa2row(8).ToString) + Len(NCa2row(9).ToString) + Len(NCa2row(19).ToString)
                                .SelectionLength = 15
                                .SelectionColor = Color.Red
                                .SelectionStart = 7 + Len(NCa2row(2)) + Len(NCa2row(7).ToString)
                                .SelectionLength = 2
                                .SelectionColor = Color.Red
                            End With
                        End If
                        If NCa2row(13) = True Then
                            nca2rt.Text += "; desp.1234/2007"
                        End If
                        If NCa2row(15) = True Then
                            nca2rt.Text += "; desp.10279/2008"
                        End If
                        If NCa2row(17) = True Then
                            nca2rt.Text += "; desp.10910/2009"
                        End If
                        If NCa2row(18) = True Then
                            nca2rt.Text += "; desp.14123/2009"
                        End If
                        If NCa2row(14) = True Then
                            nca2rt.Text += "; desp.21094/99"
                        End If
                        If NCa2row(21) = True Then
                            nca2rt.Text += "; lei 6/2010"
                        End If
                        If NCa2row(20) = False Then
                            nca2rt.SelectionStart = 0
                            nca2rt.SelectionLength = Len(NCa2row(2))
                            nca2rt.SelectionColor = Color.Purple
                        End If
                        If NCa2row(8) <> 0 And NCa2row(19) <> 0 Then
                            If NCa2row(8) > NCa2row(19) Then
                                nca2rt.SelectionStart = tamanhoantes + 10 + Len(NCa2row(2)) + Len(NCa2row(7).ToString) + Len(NCa2row(6).ToString)
                                nca2rt.SelectionLength = Len(NCa2row(8).ToString) + 4
                                nca2rt.SelectionColor = Color.Red
                            Else
                                nca2rt.SelectionStart = tamanhoantes + 10 + Len(NCa2row(2)) + Len(NCa2row(7).ToString) + Len(NCa2row(6).ToString)
                                nca2rt.SelectionLength = Len(NCa2row(8).ToString) + 4
                                nca2rt.SelectionColor = Color.Green
                            End If
                        End If
                    Else
                        With NC2rt
                            .Text = "aviado código desconhecido"
                            .ForeColor = Color.Gray
                        End With
10:                 End If
                End If
18:         Case 3
19:             If NCav3.mostrado = "true" Then
                    If Not IsNothing(NCa3row) Then
                        Dim tamanhoantes As Single
                        tamanhoantes = 0
                        'NC3dcilabel.Text = NCa3row(1) & ", " & NCa3row(4) & ", " & NCa3row(3) & ", " & NCa3row(5) & " unidade(s)"
                        If IsNumeric(NCa3row(5)) Then
                            If NCa3row(25) > 0 Then
                                NC3rt.Text = NCa3row(25) & ", " & NCa3row(1) & ", " & NCa3row(4) & ", " & NCa3row(3) & ", " & NCa3row(5) & " unidade(s)"
                            Else
                                NC3rt.Text = "semCNPEM" & ", " & NCa3row(1) & ", " & NCa3row(4) & ", " & NCa3row(3) & ", " & NCa3row(5) & " unidade(s)"
                            End If
                        Else
                            If NCa3row(25) > 0 Then
                                NC3rt.Text = NCa3row(25) & ", " & NCa3row(1) & ", " & NCa3row(4) & ", " & NCa3row(3) & ", " & NCa3row(5)
                            Else
                                NC3rt.Text = "semCNPEM" & ", " & NCa3row(1) & ", " & NCa3row(4) & ", " & NCa3row(3) & ", " & NCa3row(5)
                            End If
                        End If
                        With NC3rt
                            .SelectionStart = 0
                            .SelectionLength = 10
                            .SelectionColor = Color.Black

                            .SelectionStart = 10
                            .SelectionLength = Len(NCa3row(1)) + 2
                            .SelectionColor = Color.Blue

                            .SelectionStart = Len(NCa3row(1)) + 12
                            .SelectionLength = 2 + Len(NCa3row(4))
                            .SelectionColor = Color.Brown

                            .SelectionStart = 14 + Len(NCa3row(4)) + Len(NCa3row(1))
                            .SelectionLength = 2 + Len(NCa3row(3))
                            .SelectionColor = Color.Purple

                            .SelectionStart = 16 + Len(NCa3row(4)) + Len(NCa3row(1)) + Len(NCa3row(3))
                            .SelectionLength = Len(NCa3row(5)) + 13
                            .SelectionColor = Color.DarkGoldenrod
                        End With
                        nca3rt.Text = NCa3row(2) & "; GH0" & NCa3row(7) & "; " & NCa3row(6) & "%; PVP=" & NCa3row(8) & "; PR=" & NCa3row(9) & "; TOP5=" & NCa3row(19)
                        If NCa3row(20) = True Then
                            nca3rt.Text += "; DCI"
                        End If
                        If NCa3row(12) = True Then
                            nca3rt.Text += "; desp.13020/2011"
                            If NCa3row(20) = True Then
                                tamanhoantes = 5
                            End If
                            With nca3rt
                                .SelectionStart = tamanhoantes + 28 + Len(NCa3row(2)) + Len(NCa3row(7).ToString) + Len(NCa3row(6).ToString) + Len(NCa3row(8).ToString) + Len(NCa3row(9).ToString) + Len(NCa3row(19).ToString)
                                .SelectionLength = 15
                                .SelectionColor = Color.Red
                                .SelectionStart = 7 + Len(NCa3row(2)) + Len(NCa3row(7).ToString)
                                .SelectionLength = 2
                                .SelectionColor = Color.Red
                            End With
                        End If
                        If NCa3row(13) = True Then
                            nca3rt.Text += "; desp.1234/2007"
                        End If
                        If NCa3row(15) = True Then
                            nca3rt.Text += "; desp.10279/2008"
                        End If
                        If NCa3row(17) = True Then
                            nca3rt.Text += "; desp.10910/2009"
                        End If
                        If NCa3row(18) = True Then
                            nca3rt.Text += "; desp.14123/2009"
                        End If
                        If NCa3row(14) = True Then
                            nca3rt.Text += "; desp.21094/99"
                        End If
                        If NCa3row(21) = True Then
                            nca3rt.Text += "; lei 6/2010"
                        End If
                        If NCa3row(20) = False Then
                            nca3rt.SelectionStart = 0
                            nca3rt.SelectionLength = Len(NCa3row(2))
                            nca3rt.SelectionColor = Color.Purple
                        End If
                        If NCa3row(8) <> 0 And NCa3row(19) <> 0 Then
                            If NCa3row(8) > NCa3row(19) Then
                                nca3rt.SelectionStart = tamanhoantes + 10 + Len(NCa3row(2)) + Len(NCa3row(7).ToString) + Len(NCa3row(6).ToString)
                                nca3rt.SelectionLength = Len(NCa3row(8).ToString) + 4
                                nca3rt.SelectionColor = Color.Red
                            Else
                                nca3rt.SelectionStart = tamanhoantes + 10 + Len(NCa3row(2)) + Len(NCa3row(7).ToString) + Len(NCa3row(6).ToString)
                                nca3rt.SelectionLength = Len(NCa3row(8).ToString) + 4
                                nca3rt.SelectionColor = Color.Green
                            End If
                        End If
                    Else
                        With NC3rt
                            .Text = "aviado código desconhecido"
                            .ForeColor = Color.Gray
                        End With
                    End If
                End If
25:         Case 4
26:             If NCav4.mostrado = "true" Then
                    If Not IsNothing(NCa4row) Then
                        Dim tamanhoantes As Single
                        tamanhoantes = 0
                        'NC4dcilabel.Text = NCa4row(1) & ", " & NCa4row(4) & ", " & NCa4row(3) & ", " & NCa4row(5) & " unidade(s)"
                        If IsNumeric(NCa4row(5)) Then
                            If NCa4row(25) > 0 Then
                                NC4rt.Text = NCa4row(25) & ", " & NCa4row(1) & ", " & NCa4row(4) & ", " & NCa4row(3) & ", " & NCa4row(5) & " unidade(s)"
                            Else
                                NC4rt.Text = "semCNPEM" & ", " & NCa4row(1) & ", " & NCa4row(4) & ", " & NCa4row(3) & ", " & NCa4row(5) & " unidade(s)"
                            End If
                        Else
                            If NCa4row(25) > 0 Then
                                NC4rt.Text = NCa4row(25) & ", " & NCa4row(1) & ", " & NCa4row(4) & ", " & NCa4row(3) & ", " & NCa4row(5)
                            Else
                                NC4rt.Text = "semCNPEM" & ", " & NCa4row(1) & ", " & NCa4row(4) & ", " & NCa4row(3) & ", " & NCa4row(5)
                            End If
                        End If
                        With NC4rt
                            .SelectionStart = 0
                            .SelectionLength = 10
                            .SelectionColor = Color.Black

                            .SelectionStart = 10
                            .SelectionLength = Len(NCa4row(1)) + 2
                            .SelectionColor = Color.Blue

                            .SelectionStart = Len(NCa4row(1)) + 12
                            .SelectionLength = 2 + Len(NCa4row(4))
                            .SelectionColor = Color.Brown

                            .SelectionStart = 14 + Len(NCa4row(4)) + Len(NCa4row(1))
                            .SelectionLength = 2 + Len(NCa4row(3))
                            .SelectionColor = Color.Purple

                            .SelectionStart = 16 + Len(NCa4row(4)) + Len(NCa4row(1)) + Len(NCa4row(3))
                            .SelectionLength = Len(NCa4row(5)) + 13
                            .SelectionColor = Color.DarkGoldenrod
                        End With
                        nca4rt.Text = NCa4row(2) & "; GH0" & NCa4row(7) & "; " & NCa4row(6) & "%; PVP=" & NCa4row(8) & "; PR=" & NCa4row(9) & "; TOP5=" & NCa4row(19)
                        If NCa4row(20) = True Then
                            nca4rt.Text += "; DCI"
                        End If
                        If NCa4row(12) = True Then
                            nca4rt.Text += "; desp.13020/2011"
                            If NCa4row(20) = False Then
                                tamanhoantes = 5
                            End If
                            With nca4rt
                                'se é desp 13020/2011 avisa com despacho e comp a vermelho
                                .SelectionStart = tamanhoantes + 28 + Len(NCa4row(2)) + Len(NCa4row(7).ToString) + Len(NCa4row(6).ToString) + Len(NCa4row(8).ToString) + Len(NCa4row(9).ToString) + Len(NCa4row(19).ToString)
                                .SelectionLength = 15
                                .SelectionColor = Color.Red
                                .SelectionStart = 7 + Len(NCa4row(2)) + Len(NCa4row(7).ToString)
                                .SelectionLength = 2
                                .SelectionColor = Color.Red
                            End With
                        End If
                        If NCa4row(13) = True Then
                            nca4rt.Text += "; desp.1234/2007"
                        End If
                        If NCa4row(15) = True Then
                            nca4rt.Text += "; desp.10279/2008"
                        End If
                        If NCa4row(17) = True Then
                            nca4rt.Text += "; desp.10910/2009"
                        End If
                        If NCa4row(18) = True Then
                            nca4rt.Text += "; desp.14123/2009"
                        End If
                        If NCa4row(14) = True Then
                            nca4rt.Text += "; desp.21094/99"
                        End If
                        If NCa4row(21) = True Then
                            nca4rt.Text += "; lei 6/2010"
                        End If
                        'se não for obrigatória a prescrição por dci põe a marca a púrpura para avisar que não se pode trocar (com desp 13020/2011)
                        If NCa4row(20) = False Then
                            nca4rt.SelectionStart = 0
                            nca4rt.SelectionLength = Len(NCa4row(2))
                            nca4rt.SelectionColor = Color.Purple
                        End If
                        If NCa4row(8) <> 0 And NCa4row(19) <> 0 Then
                            If NCa4row(8) > NCa4row(19) Then
                                nca4rt.SelectionStart = tamanhoantes + 10 + Len(NCa4row(2)) + Len(NCa4row(7).ToString) + Len(NCa4row(6).ToString)
                                nca4rt.SelectionLength = Len(NCa4row(8).ToString) + 4
                                nca4rt.SelectionColor = Color.Red
                            Else
                                nca4rt.SelectionStart = tamanhoantes + 10 + Len(NCa4row(2)) + Len(NCa4row(7).ToString) + Len(NCa4row(6).ToString)
                                nca4rt.SelectionLength = Len(NCa4row(8).ToString) + 4
                                nca4rt.SelectionColor = Color.Green
                            End If
                        End If
                    Else
                        With NC4rt
                            .Text = "aviado código desconhecido"
                            .ForeColor = Color.Gray
                        End With
                    End If
                End If
32:             End Select
34:     ' End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB NCindicar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    


    Private Sub NCgb1singleclick(sender As Object, e As EventArgs) Handles NCgb1.MouseClick
        If NCaviam1.Text <> "" Then
            If NCgb1.BackColor = SystemColors.Control Then
                NCgb1.BackColor = Color.Green
                NCaviam1.BackColor = Color.Green
                nca1rt.BackColor = Color.Green
                NC1rt.BackColor = Color.Green
            ElseIf NCgb1.BackColor = Color.Red Then
                NCgb1.BackColor = SystemColors.Control
                NCaviam1.BackColor = Color.White
                nca1rt.BackColor = SystemColors.Control
                NC1rt.BackColor = SystemColors.Control
            ElseIf NCgb1.BackColor = Color.Green Then
                NCgb1.BackColor = Color.Red
                NCaviam1.BackColor = Color.Red
                nca1rt.BackColor = Color.Red
                NC1rt.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub NCgb2singleclick(sender As Object, e As EventArgs) Handles NCgb2.MouseClick
        If NCaviam2.Text <> "" Then
            If NCgb2.BackColor = SystemColors.Control Then
                NCgb2.BackColor = Color.Green
                NCaviam2.BackColor = Color.Green
                nca2rt.BackColor = Color.Green
                NC2rt.BackColor = Color.Green
            ElseIf NCgb2.BackColor = Color.Red Then
                NCgb2.BackColor = SystemColors.Control
                NCaviam2.BackColor = Color.White
                nca2rt.BackColor = SystemColors.Control
                NC2rt.BackColor = SystemColors.Control
            ElseIf NCgb2.BackColor = Color.Green Then
                NCgb2.BackColor = Color.Red
                NCaviam2.BackColor = Color.Red
                nca2rt.BackColor = Color.Red
                NC2rt.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub NCgb3singleclick(sender As Object, e As EventArgs) Handles NCgb3.MouseClick
        If NCaviam3.Text <> "" Then
            If NCgb3.BackColor = SystemColors.Control Then
                NCgb3.BackColor = Color.Green
                NCaviam3.BackColor = Color.Green
                nca3rt.BackColor = Color.Green
                NC3rt.BackColor = Color.Green
            ElseIf NCgb3.BackColor = Color.Red Then
                NCgb3.BackColor = SystemColors.Control
                NCaviam3.BackColor = Color.White
                nca3rt.BackColor = SystemColors.Control
                NC3rt.BackColor = SystemColors.Control
            ElseIf NCgb3.BackColor = Color.Green Then
                NCgb3.BackColor = Color.Red
                NCaviam3.BackColor = Color.Red
                nca3rt.BackColor = Color.Red
                NC3rt.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub NCgb4singleclick(sender As Object, e As EventArgs) Handles NCgb4.MouseClick
        If NCaviam4.Text <> "" Then
            If NCgb4.BackColor = SystemColors.Control Then
                NCgb4.BackColor = Color.Green
                NCaviam4.BackColor = Color.Green
                nca4rt.BackColor = Color.Green
                NC4rt.BackColor = Color.Green
            ElseIf NCgb4.BackColor = Color.Red Then
                NCgb4.BackColor = SystemColors.Control
                NCaviam4.BackColor = Color.White
                nca4rt.BackColor = SystemColors.Control
                NC4rt.BackColor = SystemColors.Control
            ElseIf NCgb4.BackColor = Color.Green Then
                NCgb4.BackColor = Color.Red
                NCaviam4.BackColor = Color.Red
                nca4rt.BackColor = Color.Red
                NC4rt.BackColor = Color.Red
            End If
        End If
    End Sub
End Class