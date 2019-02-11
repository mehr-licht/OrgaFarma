Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb

Public Class Form3
    Inherits Form

    Dim foco3 As String

    Dim A As Short
    Dim av1row As basededadosDataSet.infarmedRow
    Dim av2row As basededadosDataSet.infarmedRow
    Dim av3row As basededadosDataSet.infarmedRow
    Dim av4row As basededadosDataSet.infarmedRow
    Dim av5row As basededadosDataSet.infarmedRow
    Dim av6row As basededadosDataSet.infarmedRow
    Dim codigorow As basededadosDataSet.infarmedRow
    Dim av1array As New ArrayList
    Dim av2array As New ArrayList
    Dim av3array As New ArrayList
    Dim av4array As New ArrayList
    Dim av5array As New ArrayList
    Dim av6array As New ArrayList
    Dim Aviad1 As New meds
    Dim Aviad2 As New meds
    Dim Aviad3 As New meds
    Dim Aviad4 As New meds
    Dim Aviad5 As New meds
    Dim Aviad6 As New meds
    Dim vazio3 As New meds
    Dim gen As Boolean
    Dim infarmedTA3 As New basededadosDataSetTableAdapters.infarmedTableAdapter
    Dim DS3 As New basededadosDataSet
    Dim row3 As DataRow
    Dim comp As Double
    'esta variavel guarda o número de linhas lido da base de dados
    Dim infarmedTA3s As Integer = infarmedTA3.Fill(DS3.infarmed)

    Dim org As Short
    Dim grupo3P1 As Short = 0
    Dim grupo3P2 As Short = 0
    Dim grupo3P3 As Short = 0
    Dim grupo3P4 As Short = 0
    Dim grupo3A1 As Short = 0
    Dim grupo3A2 As Short = 0
    Dim grupo3A3 As Short = 0
    Dim grupo3A4 As Short = 0
    Dim grupo3P1dci As String
    Dim grupo3P2dci As String
    Dim grupo3P3dci As String
    Dim grupo3P4dci As String
    Dim grupo3A1dci As String
    Dim grupo3A2dci As String
    Dim grupo3A3dci As String
    Dim grupo3A4dci As String
    Dim qualport As Integer
    Dim organismo As Short
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
    Dim portcomp As Double
    Dim portcomp1 As Double
    Dim portcomp2 As Double
    Dim portcomp3 As Double
    Dim portcomp4 As Double
    Dim portcomp5 As Double
    Dim portcomp6 As Double
    Dim intermedio As Double
    Dim pr As Double
    Dim pr1 As Double
    Dim pr2 As Double
    Dim pr3 As Double
    Dim pr4 As Double
    Dim pr5 As Double
    Dim pr6 As Double
    Dim portimedio As String
    Dim avio1 As New avaliacao
    Dim avio2 As New avaliacao
    Dim avio3 As New avaliacao
    Dim avio4 As New avaliacao
    Dim avio5 As New avaliacao
    Dim avio6 As New avaliacao
    Dim grupo1dci As String
    Dim grupo2dci As String
    Dim grupo3dci As String
    Dim grupo4dci As String
    Dim grupo5dci As String
    Dim grupo6dci As String



    Sub agrupar3()
        On Error GoTo MOSTRARERRO


        If A >= 1 Then
            If Not IsNothing(av1row) Then
                grupo1dci = av1row(1).ToString
                avio1.grupo = 1
            End If
        End If


        If A >= 2 Then
            If Not IsNothing(av2row) Then
                If av1row(1) = av2row(1) And av1row(2) = av2row(2) And av1row(3) = av2row(3) Then
                    avio2.grupo = 1
                Else
                    grupo2dci = av2row(1).ToString
                    avio2.grupo = 2
                End If
            End If
        End If

        If A >= 3 Then
            If Not IsNothing(av3row) Then
                If av1row(1) = av3row(1) And av1row(2) = av3row(2) And av1row(3) = av3row(3) Then
                    avio3.grupo = 1
                ElseIf av2row(1) = av3row(1) And av2row(2) = av3row(2) And av2row(3) = av3row(3) Then
                    avio3.grupo = avio2.grupo
                Else
                    grupo3dci = av3row(1).ToString
                    avio3.grupo = 3
                End If
            End If
        End If

        If A >= 4 Then
            If Not IsNothing(av4row) Then
                If av1row(1) = av4row(1) And av1row(2) = av4row(2) And av1row(3) = av4row(3) Then
                    avio4.grupo = 1
                ElseIf av2row(1) = av4row(1) And av2row(2) = av4row(2) And av2row(3) = av4row(3) Then
                    avio4.grupo = avio2.grupo
                ElseIf av3row(1) = av4row(1) And av3row(2) = av4row(2) And av3row(3) = av4row(3) Then
                    avio4.grupo = avio3.grupo
                Else
                    grupo4dci = av4row(1).ToString
                    avio4.grupo = 4
                End If
            End If
        End If

        If A >= 5 Then
            If Not IsNothing(av5row) Then
                If av1row(1) = av5row(1) And av1row(2) = av5row(2) And av1row(3) = av5row(3) Then
                    avio5.grupo = 1
                ElseIf av2row(1) = av5row(1) And av2row(2) = av5row(2) And av2row(3) = av5row(3) Then
                    avio5.grupo = avio2.grupo
                ElseIf av3row(1) = av5row(1) And av3row(2) = av5row(2) And av3row(3) = av5row(3) Then
                    avio5.grupo = avio3.grupo
                ElseIf av4row(1) = av5row(1) And av4row(2) = av5row(2) And av4row(3) = av5row(3) Then
                    avio5.grupo = avio4.grupo
                Else
                    grupo5dci = av5row(1).ToString
                    avio5.grupo = 5
                End If
            End If
        End If

        If A >= 6 Then
            If Not IsNothing(av6row) Then
                If av1row(1) = av6row(1) And av1row(2) = av6row(2) And av1row(3) = av6row(3) Then
                    avio6.grupo = 1
                ElseIf av2row(1) = av6row(1) And av2row(2) = av6row(2) And av2row(3) = av6row(3) Then
                    avio6.grupo = avio2.grupo
                ElseIf av3row(1) = av6row(1) And av3row(2) = av6row(2) And av3row(3) = av6row(3) Then
                    avio6.grupo = avio3.grupo
                ElseIf av4row(1) = av6row(1) And av4row(2) = av6row(2) And av4row(3) = av6row(3) Then
                    avio6.grupo = avio4.grupo
                ElseIf av5row(1) = av6row(1) And av5row(2) = av6row(2) And av5row(3) = av6row(3) Then
                    avio6.grupo = avio5.grupo
                Else
                    grupo6dci = av6row(1).ToString
                    avio6.grupo = 6
                End If
            End If
        End If

        If avio1.mostrado Then
            If avio1.grupo = avio2.grupo And avio1.grupo = avio3.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av2.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio2.grupo And avio1.grupo = avio4.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av2.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio2.grupo And avio1.grupo = avio5.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av2.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio2.grupo And avio1.grupo = avio6.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av2.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio3.grupo And avio1.grupo = avio4.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio3.grupo And avio1.grupo = avio5.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio3.grupo And avio1.grupo = avio6.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio4.grupo And avio1.grupo = avio5.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio1.grupo = avio5.grupo And avio1.grupo = avio6.grupo Then
                av1.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
        End If
        If avio2.mostrado Then
            If avio2.grupo = avio3.grupo And avio2.grupo = avio4.grupo Then
                av2.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio2.grupo = avio3.grupo And avio2.grupo = avio5.grupo Then
                av2.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio2.grupo = avio3.grupo And avio2.grupo = avio6.grupo Then
                av2.BackColor = Color.DarkGoldenrod
                av3.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio2.grupo = avio4.grupo And avio2.grupo = avio5.grupo Then
                av2.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio2.grupo = avio4.grupo And avio2.grupo = avio6.grupo Then
                av2.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
        End If
        If avio3.mostrado Then
            If avio3.grupo = avio4.grupo And avio3.grupo = avio5.grupo Then
                av3.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio3.grupo = avio4.grupo And avio3.grupo = avio6.grupo Then
                av3.BackColor = Color.DarkGoldenrod
                av4.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
            If avio3.grupo = avio5.grupo And avio3.grupo = avio6.grupo Then
                av3.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
        End If
        If avio4.mostrado Then
            If avio4.grupo = avio5.grupo And avio4.grupo = avio6.grupo Then
                av4.BackColor = Color.DarkGoldenrod
                av5.BackColor = Color.DarkGoldenrod
                av6.BackColor = Color.DarkGoldenrod
                MsgBox("aviados mais de dois medicamentos iguais")
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub start()
        'não muda o organismo (continua o mesmo que já estava
        On Error GoTo MOSTRARERRO
        labelatribuido.Text = ""
        'Me.pvp1val = 0
        'Me.pvp2val = 0
        'Me.pvp3val = 0
        'Me.pvp4val = 0
        'Me.pvp5val = 0
        'Me.pvp6val = 0
        'Me.pvp1v = ""
        'Me.pvp2v = ""
        'Me.pvp3v = ""
        'Me.pvp4v = ""
        'Me.pvp5v = ""
        'Me.pvp6v = ""
        Me.av1.Text = ""
        Me.av2.Text = ""
        Me.av3.Text = ""
        Me.av4.Text = ""
        Me.av5.Text = ""
        Me.av6.Text = ""
        Me.gen1.Text = ""
        Me.gen2.Text = ""
        Me.gen3.Text = ""
        Me.gen4.Text = ""
        Me.gen5.Text = ""
        Me.gen6.Text = ""
        Me.port1.Text = ""
        Me.port2.Text = ""
        Me.port3.Text = ""
        Me.port4.Text = ""
        Me.port5.Text = ""
        Me.port6.Text = ""
        Me.pvp1.Text = ""
        Me.pvp2.Text = ""
        Me.pvp3.Text = ""
        Me.pvp4.Text = ""
        Me.pvp5.Text = ""
        Me.pvp6.Text = ""
        Me.comp1.Text = ""
        Me.comp2.Text = ""
        Me.comp3.Text = ""
        Me.comp4.Text = ""
        Me.comp5.Text = ""
        Me.comp6.Text = ""
        Me.totalPVP.Text = ""
        Me.totalComp.Text = ""
        tirarports()
        Aviad1 = vazio3
        Aviad2 = vazio3
        Aviad3 = vazio3
        Aviad4 = vazio3
        Aviad5 = vazio3
        Aviad6 = vazio3
        av1.BackColor = Color.White
        av2.BackColor = Color.White
        av3.BackColor = Color.White
        av4.BackColor = Color.White
        av5.BackColor = Color.White
        av6.BackColor = Color.White
        ' grupo3P1 = 0
        'grupo3P2 = 0
        'grupo3P3 = 0
        'grupo3P4 = 0
        'grupo3A1 = 0
        'grupo3A2 = 0
        'grupo3A3 = 0
        'grupo3A4 = 0
        'grupo3P1dci = ""
        'grupo3P2dci = ""
        'grupo3P3dci = ""
        'grupo3P4dci = ""
        'grupo3A1dci = ""
        'grupo3A2dci = ""
        'grupo3A3dci = ""
        'grupo3A4dci = ""
        If IsNothing(gen) Then
            gen = "false"
        End If
        If IsNothing(portcomp) Then
            portcomp = 0
        End If
        If IsNothing(organismo) Then
            organismo = 1
        End If
        av1.Focus()
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB start: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'não está a ser usado
    Sub irbuscar3(ByVal qualwhich As Short)
        On Error GoTo MOSTRARERRO

        Exit Sub
MOSTRARERRO:
        MsgBox("SUB irbsucar3: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'valida se só contém algarismos. se não dá beep e fica onde está. se sim verifica se tem 7 algarismos e passa à frente. o que acontece ao fazer enter
    Private Sub frmDesigner_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        On Error GoTo MOSTRARERRO
        foco3 = Me.ActiveControl.Name()
        Dim soalgarismos As Boolean
        If Not Char.IsNumber(e.KeyChar) Then
            soalgarismos = False
        Else
            soalgarismos = True
        End If

        Select Case soalgarismos
            Case False
                Beep()
        End Select




        If Asc(e.KeyChar) = Keys.Enter Then
            Select Case foco3
                Case "av1"
                    If av1.Text = "" Then
                        Beep()
                    Else
                        Me.av2.Focus()
                    End If

                Case "av2"
                    If av2.Text = "" Then
                        somas()
                    Else
                        Me.av3.Focus()
                    End If
                Case "av3"
                    If av3.Text = "" Then
                        somas()
                    Else
                        Me.av4.Focus()
                    End If
                Case "av4"
                    If av4.Text = "" Then
                        somas()
                    Else
                        Me.av5.Focus()
                    End If
                Case "av5"
                    If av5.Text = "" Then
                        somas()
                    Else
                        Me.av6.Focus()
                    End If
                Case "av6"
                    If av6.Text = "" Then
                        somas()
                    Else
                        If av6.Text >= 1111111 And av6.Text <= 9999999 Then
                            Me.pvp1.Focus()
                            somas()
                        Else
                            Beep()
                        End If
                    End If
            End Select
        End If


        If Asc(e.KeyChar) = Keys.C Then
            My.Computer.Keyboard.SendKeys("{bs}")
            My.Computer.Keyboard.SendKeys("{bs}")
            My.Computer.Keyboard.SendKeys("{ENTER}")
        End If



        If Asc(e.KeyChar) = Keys.Space Then
            start()
        End If



        Exit Sub
MOSTRARERRO:
        MsgBox("sub keypress: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'os próximos 6 são os validadores que estão a funcionar. faziam (com caneta não fazem) saltar o focus quando se inserem 7 caracteres. e lançam a comparação
    Private Sub av1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av1.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta1 As String
        Caracta1 = av1.Text
        If Len(av1.Text) = 7 Then
            If Caracta1 Like "#######" Then
                Aviad1.codigo = av1.Text
                avio1.mostrado = "True"
                A = 1
                av1row = DS3.infarmed.FindBycode(Aviad1.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad1.codigo)
                av1array.Add(av1row)
                indicar(1)
                'Me.av2.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av1_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub av2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av2.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta2 As String
        Caracta2 = av2.Text
        If Len(av2.Text) = 7 Then
            If Caracta2 Like "#######" Then
                Aviad2.codigo = av2.Text
                avio2.mostrado = "True"
                A = 2
                av2row = DS3.infarmed.FindBycode(Aviad2.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad2.codigo)
                av2array.Add(av2row)
                indicar(2)
                'Me.av3.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av2_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub av3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av3.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta3 As String
        Caracta3 = av3.Text
        If Len(av3.Text) = 7 Then
            If Caracta3 Like "#######" Then
                Aviad3.codigo = av3.Text
                avio3.mostrado = "True"
                A = 3
                av3row = DS3.infarmed.FindBycode(Aviad3.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad3.codigo)
                av3array.Add(av3row)
                indicar(3)
                'Me.av4.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av3_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub av4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av4.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta4 As String
        Caracta4 = av4.Text
        If Len(av4.Text) = 7 Then
            If Caracta4 Like "#######" Then
                Aviad4.codigo = av4.Text
                avio4.mostrado = "True"
                A = 4
                av4row = DS3.infarmed.FindBycode(Aviad4.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad4.codigo)
                av4array.Add(av4row)
                indicar(4)
                'Me.av5.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av4_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub av5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av5.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta5 As String
        Caracta5 = av5.Text
        If Len(av5.Text) = 7 Then
            If Caracta5 Like "#######" Then
                Aviad5.codigo = av5.Text
                avio5.mostrado = "True"
                A = 5
                av5row = DS3.infarmed.FindBycode(Aviad5.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad5.codigo)
                av5array.Add(av5row)
                indicar(5)
                'Me.av6.Focus()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av5_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub av6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles av6.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta6 As String
        Caracta6 = av6.Text
        If Len(av6.Text) = 7 Then
            If Caracta6 Like "#######" Then
                Aviad6.codigo = av6.Text
                avio6.mostrado = "True"
                A = 6
                av6row = DS3.infarmed.FindBycode(Aviad6.codigo)
                codigorow = DS3.infarmed.FindBycode(Aviad6.codigo)
                av6array.Add(av6row)
                indicar(6)
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("av6_TextChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub indicar(ByVal which As Short)
        On Error GoTo MOSTRARERRO
        If Not IsNothing(codigorow) Then
            Select Case which
                Case 1
                    If avio1.mostrado = "true" Then
                        If av1row(7) = "true" Then
                            gen1.Text = "genérico"
                            gen = True
                        Else
                            gen1.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port1.Text = portimedio
                        comp = (av1row(5) * 0.01)
                        portcomp1 = portcomp
                        intermedio = Replace(av1row(17), ".", ",")
                        pvp1.Text = intermedio
                        pr1 = Replace(av1row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr1 = 1.2 * pr1
                        End If
                        pr = pr1
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp1.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
                Case 2
                    If avio2.mostrado = "true" Then
                        If av2row(7) = "true" Then
                            gen2.Text = "genérico"
                            gen = True
                        Else
                            gen2.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port2.Text = portimedio
                        comp = (av2row(5) * 0.01)
                        portcomp2 = portcomp
                        intermedio = Replace(av2row(17), ".", ",")
                        pvp2.Text = intermedio
                        pr2 = Replace(av2row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr2 = 1.2 * pr2
                        End If
                        pr = pr2
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp2.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
                Case 3
                    If avio3.mostrado = "true" Then
                        If av3row(7) = "true" Then
                            gen3.Text = "genérico"
                            gen = True
                        Else
                            gen3.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port3.Text = portimedio
                        comp = (av3row(5) * 0.01)
                        portcomp3 = portcomp
                        intermedio = Replace(av3row(17), ".", ",")
                        pvp3.Text = intermedio
                        pr3 = Replace(av3row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr3 = 1.2 * pr3
                        End If
                        pr = pr3
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp3.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
                Case 4
                    If avio4.mostrado = "true" Then
                        If av4row(7) = "true" Then
                            gen4.Text = "genérico"
                            gen = True
                        Else
                            gen4.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port4.Text = portimedio
                        comp = (av4row(5) * 0.01)
                        portcomp4 = portcomp
                        intermedio = Replace(av4row(17), ".", ",")
                        pvp4.Text = intermedio
                        pr4 = Replace(av4row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr4 = 1.2 * pr4
                        End If
                        pr = pr4
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp4.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
                Case 5
                    If avio5.mostrado = "true" Then
                        If av5row(7) = "true" Then
                            gen5.Text = "genérico"
                            gen = True
                        Else
                            gen5.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port5.Text = portimedio
                        comp = (av5row(5) * 0.01)
                        portcomp5 = portcomp
                        intermedio = Replace(av5row(17), ".", ",")
                        pvp5.Text = intermedio
                        pr5 = Replace(av5row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr5 = 1.2 * pr5
                        End If
                        pr = pr5
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp5.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
                Case 6
                    If avio6.mostrado = "true" Then
                        If av6row(7) = "true" Then
                            gen6.Text = "genérico"
                            gen = True
                        Else
                            gen6.Text = "marca"
                            gen = False
                        End If
                        portaria()
                        port6.Text = portimedio
                        comp = (av6row(5) * 0.01)
                        portcomp6 = portcomp
                        intermedio = Replace(av6row(17), ".", ",")
                        pvp6.Text = intermedio
                        pr6 = Replace(av6row(18), ".", ",")
                        If organismo = 48 Or organismo = 49 Then
                            pr6 = 1.2 * pr6
                        End If
                        pr = pr6
                        If pr > 0 Then
                            intermedio = System.Math.Min(intermedio, pr)
                        End If
                        comp6.Text = System.Math.Round(calculo(organismo, gen, comp, intermedio), 2)
                    End If
            End Select
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB indicar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'não está a ser usado
    Public Sub atribuira3()
        On Error GoTo MOSTRARERRO
        If av6.Text <> "0" And av6.Text <> "" Then
            A = 6
        ElseIf av5.Text <> "0" And av5.Text <> "" Then
            A = 5
        ElseIf av4.Text <> "0" And av4.Text <> "" Then
            A = 4
        ElseIf av3.Text <> "0" And av3.Text <> "" Then
            A = 3
        ElseIf av2.Text <> "0" And av2.Text <> "" Then
            A = 2
        Else : A = 1
        End If


        If A = 6 Then
            Aviad6.principio = av6row(1)
            Aviad6.apresentacao = av6row(2)
            Aviad6.dosagem = av6row(3)
            Aviad6.quantidade = av6row(4)
            Aviad6.comparticipacao = av6row(5)
            Aviad6.grupo = av6row(6)
            Aviad6.generico = av6row(7)
            Aviad6.laboratorio = av6row(8)
        ElseIf A >= 5 Then
            If A = 5 Then
                Aviad6 = vazio3
            End If
            Aviad5.principio = av5row(1)
            Aviad5.apresentacao = av5row(2)
            Aviad5.dosagem = av5row(3)
            Aviad5.quantidade = av5row(4)
            Aviad5.comparticipacao = av5row(5)
            Aviad5.grupo = av5row(6)
            Aviad5.generico = av5row(7)
            Aviad5.laboratorio = av5row(8)
        ElseIf A >= 4 Then
            If A = 4 Then
                Aviad6 = vazio3
                Aviad5 = vazio3
            End If
            Aviad4.principio = av4row(1)
            Aviad4.apresentacao = av4row(2)
            Aviad4.dosagem = av4row(3)
            Aviad4.quantidade = av4row(4)
            Aviad4.comparticipacao = av4row(5)
            Aviad4.grupo = av4row(6)
            Aviad4.generico = av4row(7)
            Aviad4.laboratorio = av4row(8)
        ElseIf A >= 3 Then
            If A = 3 Then
                Aviad6 = vazio3
                Aviad5 = vazio3
                Aviad4 = vazio3
            End If
            Aviad3.principio = av3row(1)
            Aviad3.apresentacao = av3row(2)
            Aviad3.dosagem = av3row(3)
            Aviad3.quantidade = av3row(4)
            Aviad3.comparticipacao = av3row(5)
            Aviad3.grupo = av3row(6)
            Aviad3.generico = av3row(7)
            Aviad3.laboratorio = av3row(8)
        ElseIf A >= 2 Then
            If A = 2 Then
                Aviad6 = vazio3
                Aviad5 = vazio3
                Aviad4 = vazio3
                Aviad3 = vazio3
            End If
            Aviad2.principio = av2row(1)
            Aviad2.apresentacao = av2row(2)
            Aviad2.dosagem = av2row(3)
            Aviad2.quantidade = av2row(4)
            Aviad2.comparticipacao = av2row(5)
            Aviad2.grupo = av2row(6)
            Aviad2.generico = av2row(7)
            Aviad2.laboratorio = av2row(8)
        ElseIf A >= 1 Then
            If A = 1 Then
                Aviad6 = vazio3
                Aviad5 = vazio3
                Aviad4 = vazio3
                Aviad3 = vazio3
                Aviad2 = vazio3
            End If
            Aviad1.principio = av1row(1)
            Aviad1.apresentacao = av1row(2)
            Aviad1.dosagem = av1row(3)
            Aviad1.quantidade = av1row(4)
            Aviad1.comparticipacao = av1row(5)
            Aviad1.grupo = av1row(6)
            Aviad1.generico = av1row(7)
            Aviad1.laboratorio = av1row(8)
            End If
            Exit Sub
MOSTRARERRO:
            MsgBox("sub atribuira3: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
            Resume Next
    End Sub

    Function SomarPVP()
        On Error GoTo MOSTRARERRO
        Dim somaPVP As Double
        If pvp1.Text <> "" Then
            pvp1v = Replace(pvp1.Text, ".", ",")
            pvp1val = Convert.ToDouble(pvp1v)
        Else : pvp1val = 0
        End If
        If pvp2.Text <> "" Then
            pvp2v = Replace(pvp2.Text, ".", ",")
            pvp2val = Convert.ToDouble(pvp2v)
        Else : pvp2val = 0
        End If
        If pvp3.Text <> "" Then
            pvp3v = Replace(pvp3.Text, ".", ",")
            pvp3val = Convert.ToDouble(pvp3v)
        Else : pvp3val = 0
        End If
        If pvp4.Text <> "" Then
            pvp4v = Replace(pvp4.Text, ".", ",")
            pvp4val = Convert.ToDouble(pvp4v)
        Else : pvp4val = 0
        End If
        If pvp5.Text <> "" Then
            pvp5v = Replace(pvp5.Text, ".", ",")
            pvp5val = Convert.ToDouble(pvp5v)
        Else : pvp5val = 0
        End If
        If pvp6.Text <> "" Then
            pvp6v = Replace(pvp6.Text, ".", ",")
            pvp6val = Convert.ToDouble(pvp6v)
        Else : pvp6val = 0
        End If
        somaPVP = pvp1val + pvp2val + pvp3val + pvp4val + pvp5val + pvp6val
        SomarPVP = somaPVP
        Exit Function
MOSTRARERRO:
        MsgBox("SUB somarpvp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Function SomarComp()
        On Error GoTo MOSTRARERRO
        Dim somaComp As Double
        If comp1.Text <> "" Then
            comp1v = Replace(comp1.Text, ".", ",")
            comp1val = Convert.ToDouble(comp1v)
        Else : comp1val = 0
        End If
        If comp2.Text <> "" Then
            comp2v = Replace(comp2.Text, ".", ",")
            comp2val = Convert.ToDouble(comp2v)
        Else : comp2val = 0
        End If
        If comp3.Text <> "" Then
            comp3v = Replace(comp3.Text, ".", ",")
            comp3val = Convert.ToDouble(comp3v)
        Else : comp3val = 0
        End If
        If comp4.Text <> "" Then
            comp4v = Replace(comp4.Text, ".", ",")
            comp4val = Convert.ToDouble(comp4v)
        Else : comp4val = 0
        End If
        If comp5.Text <> "" Then
            comp5v = Replace(comp5.Text, ".", ",")
            comp5val = Convert.ToDouble(comp5v)
        Else : comp5val = 0
        End If
        If comp6.Text <> "" Then
            comp6v = Replace(comp6.Text, ".", ",")
            comp6val = Convert.ToDouble(comp6v)
        Else : comp6val = 0
        End If
        somaComp = comp1val + comp2val + comp3val + comp4val + comp5val + comp6val
        SomarComp = somaComp
        Exit Function
MOSTRARERRO:
        MsgBox("SUB somarcomp: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function

    Sub somas()
        On Error GoTo MOSTRARERRO
        agrupar3()
        SomarPVP()
        SomarComp()
        totalPVP.Text = SomarPVP()
        totalComp.Text = SomarComp()
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB somas: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo MOSTRARERRO
        Me.KeyPreview = True
        start()
        Exit Sub
MOSTRARERRO:
        MsgBox("sub form calcular load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub butStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butStart.Click
        On Error GoTo MOSTRARERRO
        start()
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB butStart_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub butCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCalc.Click
        On Error GoTo MOSTRARERRO
        somas()
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB butCalc_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub portaria()
        On Error GoTo MOSTRARERRO
        If codigorow(9) = True Then
            portimedio = "4250"
        End If

        If codigorow(10) = True Then
            portimedio = "1234"
        End If

        If codigorow(11) = True Then
            portimedio = "10279"
        End If

        If codigorow(12) = True Then
            portimedio = "10280"
        End If

        If codigorow(13) = True Then
            portimedio = "10910"
        End If

        If codigorow(14) = True Then
            If codigorow(10) = True Then
                portimedio = "1234 + 14123"
            Else
                portimedio = "14123"
            End If
        End If

        If codigorow(15) = True Then
            portimedio = "147469"
        End If

        If codigorow(19) = True Then
            portimedio = "21094"
        End If

        If codigorow(20) = True Then
            portimedio = "1474100"
        End If
        If codigorow(9) = False And codigorow(10) = False And codigorow(11) = False And codigorow(12) = False And codigorow(13) = False _
         And codigorow(14) = False And codigorow(15) = False And codigorow(19) = False And codigorow(20) = False Then
            portimedio = "não"
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB portaria: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'já não é usado
    Sub DesComPort(ByVal escolher As String)
        On Error GoTo MOSTRARERRO
        Select Case escolher
            Case 4250
                portcomp = Replace(0.37, ".", ",")
            Case 1234
                portcomp = Replace(0.95, ".", ",")
            Case 10279
                portcomp = Replace(0.95, ".", ",")
            Case 10280
                portcomp = Replace(0.95, ".", ",")
            Case 10910
                portcomp = Replace(0.69, ".", ",")
            Case 14123
                portcomp = Replace(0.69, ".", ",")
            Case 147469
                portcomp = Replace(0.69, ".", ",")
            Case 1474100
                portcomp = 1
            Case 21094
                portcomp = 1

        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB DesComPort_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next

    End Sub


    Function calculo(ByVal org As Short, ByVal gen As Boolean, ByVal comp As Double, ByVal intermedio As Double) As Double
        On Error GoTo MOSTRARERRO

        Select Case org
            Case 1 'tipo 10
                calculo = intermedio * comp
            Case 46 'tipo 17
                calculo = intermedio * comp
            Case 42 'tipo 12
                calculo = intermedio
            Case 41 'tipo 11
                If comp > 0 Then
                    calculo = intermedio
                Else
                    calculo = 0
                End If
            Case 67 'tipo 13
                If comp > 0 Then
                    calculo = intermedio
                Else
                    calculo = 0
                End If
            Case 23, 24, 25 'diabetes - não sei se 24 e 25 também são assim mas já fica
                calculo = intermedio * 0.85
            Case 48 'tipo 15
                If comp > 0 Then
                    If gen = "true" Then
                        calculo = intermedio
                    Else
                        calculo = (System.Math.Min(1, (comp + 0.15))) * intermedio
                    End If
                Else
                    calculo = 0
                End If
            Case 45
                calculo = intermedio * (System.Math.Max(comp, portcomp))
            Case 49
                If gen = "true" Then
                    calculo = intermedio
                Else
                    calculo = System.Math.Min(1, (System.Math.Max((portcomp + 0.15), (comp + 0.15)))) * intermedio
                End If
        End Select

        'não tenho nada para o tipo 19 nem para os organismos 02_ADSE,12_BancNt,15_IASFA,17_GNR,18_PSP,25_SAMSq,59_ADSEdipl,R5_CGD

        Exit Function
MOSTRARERRO:
        MsgBox("SUB calculo: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function



    Private Sub but1474_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_01.Checked Then
            portimedio = port1.Text
            If av1row(15) = "true" And av1row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av1row(20) = "true" And av1row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(1)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_02.Checked Then
            portimedio = port2.Text
            If av2row(15) = "true" And av2row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av2row(20) = "true" And av2row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(2)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_03.Checked Then
            portimedio = port3.Text
            If av3row(15) = "true" And av3row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av3row(20) = "true" And av3row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(3)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_04.Checked Then
            If av4row(15) = "true" And av4row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av4row(20) = "true" And av4row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(4)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_05.Checked Then
            portimedio = port5.Text
            If av5row(15) = "true" And av5row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av5row(20) = "true" And av5row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(5)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1474_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1474_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1474_06.Checked Then
            portimedio = port6.Text
            If av6row(15) = "true" And av6row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            ElseIf av6row(20) = "true" And av6row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            indicar(6)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1474_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_01.Checked Then
            portimedio = port1.Text
            If av1row(10) = "true" And av1row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(1)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_02.Checked Then
            portimedio = port2.Text
            If av2row(10) = "true" And av2row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(2)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_03.Checked Then
            portimedio = port3.Text
            If av3row(10) = "true" And av3row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(3)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_04.Checked Then
            portimedio = port4.Text
            If av4row(10) = "true" And av4row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(4)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_05.Checked Then
            portimedio = port5.Text
            If av5row(10) = "true" And av5row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(5)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but1234_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but1234_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but1234_06.Checked Then
            portimedio = port6.Text
            If av6row(10) = "true" And av6row(5) <> 0 Then
                portcomp = Replace(0.95, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(6)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but1234_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_01.Checked Then
            portimedio = port1.Text
            If av1row(11) = "true" Or av1row(12) = "true" Then
                If av1row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(1)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_02.Checked Then
            portimedio = port2.Text
            If av2row(11) = "true" Or av2row(12) = "true" Then
                If av2row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(2)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but10279_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_03.Checked Then
            portimedio = port3.Text
            If av3row(11) = "true" Or av3row(12) = "true" Then
                If av3row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(3)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_04.Checked Then
            portimedio = port4.Text
            If av4row(11) = "true" Or av4row(12) = "true" Then
                If av4row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(4)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10279_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_05.Checked Then
            portimedio = port5.Text
            If av5row(11) = "true" Or av5row(12) = "true" Then
                If av5row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(5)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub but10279_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10279_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10279_06.Checked Then
            portimedio = port6.Text
            If av6row(11) = "true" Or av6row(12) = "true" Then
                If av6row(5) <> 0 Then
                    portcomp = Replace(0.95, ".", ",")
                Else
                    portcomp = 0
                End If
                indicar(6)
                somas()
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10279_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but14123_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_01.Checked Then
            portimedio = port1.Text
            If av1row(14) = "true" And av1row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(1)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_02.Checked Then
            portimedio = port2.Text
            If av2row(14) = "true" And av2row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(2)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_03.Checked Then
            portimedio = port3.Text
            If av3row(14) = "true" And av3row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(3)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_04.Checked Then
            portimedio = port4.Text
            If av4row(14) = "true" And av4row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(4)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_05.Checked Then
            portimedio = port5.Text
            If av5row(14) = "true" And av5row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(5)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but14123_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but14123_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but14123_06.Checked Then
            portimedio = port6.Text
            If av6row(14) = "true" And av6row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(6)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but14123_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_01.Checked Then
            portimedio = port1.Text
            If av1row(13) = "true" And av1row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(1)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_02.Checked Then
            portimedio = port2.Text
            If av2row(13) = "true" And av2row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(2)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_03.Checked Then
            portimedio = port3.Text
            If av3row(13) = "true" And av3row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(3)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_04.Checked Then
            portimedio = port4.Text
            If av4row(13) = "true" And av4row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(4)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_05.Checked Then
            portimedio = port5.Text
            If av5row(13) = "true" And av5row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(5)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but10910_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but10910_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but10910_06.Checked Then
            portimedio = port6.Text
            If av6row(13) = "true" And av6row(5) <> 0 Then
                portcomp = Replace(0.69, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(6)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but10910_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_01.Checked Then
            portimedio = port1.Text
            If av1row(19) = "true" And av1row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_02.Checked Then
            portimedio = port2.Text
            If av2row(19) = "true" And av2row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_03.Checked Then
            portimedio = port3.Text
            If av3row(19) = "true" And av3row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but21094_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_04.Checked Then
            portimedio = port4.Text
            If av4row(19) = "true" And av4row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_05.Checked Then
            portimedio = port5.Text
            If av5row(19) = "true" And av5row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but21094_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but21094_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but21094_06.Checked Then
            portimedio = port6.Text
            If av6row(19) = "true" And av6row(5) <> 0 Then
                portcomp = 1
            Else
                portcomp = 0
            End If
            For i = 1 To 6
                indicar(i)
            Next
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but21094_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_01.Checked Then
            portimedio = port1.Text
            If av1row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(1)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_02_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_02.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_02.Checked Then
            portimedio = port2.Text
            If av2row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(2)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_02_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_03_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_03.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_03.Checked Then
            portimedio = port3.Text
            If av3row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(3)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_03_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_04_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_04.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_04.Checked Then
            portimedio = port4.Text
            If av4row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(4)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_04_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_05_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_05.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_05.Checked Then
            portimedio = port5.Text
            If av5row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(5)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_05_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but4250_06_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but4250_06.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but4250_06.Checked Then
            portimedio = port6.Text
            If av6row(9) = "true" Then
                portcomp = Replace(0.37, ".", ",")
            Else
                portcomp = 0
            End If
            indicar(6)
            somas()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but4250_06_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but01.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but01.Checked Then
            deslabelar()
            organismo = 1
            organismus(organismo)
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but48_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but48.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but48.Checked Then
            deslabelar()
            organismo = 48
            organismus(organismo)
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but01_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but41_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but41.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but41.Checked Then
            deslabelar()
            organismo = 41
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "doentes profissionais"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but41_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but46_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but46.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but46.Checked Then
            deslabelar()
            organismo = 46
            organismus(organismo)
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but46_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but42_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but42.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but42.Checked Then
            deslabelar()
            organismo = 42
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "portaria 4521/2001"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but42_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but67_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but67.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but67.Checked Then
            deslabelar()
            organismo = 67
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "despacho 11387-A/2003"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but67_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub butDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDS.CheckedChanged
        On Error GoTo MOSTRARERRO
        If butDS.Checked Then
            deslabelar()
            organismo = 23
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "diabetes SNS"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB butDS_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Private Sub but49_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but49.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but49.Checked Then
            deslabelar()
            organismo = 49
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "portaria 1474/2004"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but49_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but45_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but45.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but45.Checked Then
            deslabelar()
            organismo = 45
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "portaria 1474/2004"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but45_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub but59_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but59.CheckedChanged
        On Error GoTo MOSTRARERRO
        If but59.Checked Then
            deslabelar()
            organismo = 59
            organismus(organismo)
            labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Bold)
            labelatribuido.Text = "portaria 1474/2004"
            av1.Focus()
        End If
        For i = 1 To 6
            indicar(i)
            somas()
        Next
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB but59_CheckedChanged: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub tirarports()
        On Error GoTo MOSTRARERRO
        but1474_01.Checked = False
        but1474_02.Checked = False
        but1474_03.Checked = False
        but1474_04.Checked = False
        but1474_05.Checked = False
        but1474_06.Checked = False
        but1234_01.Checked = False
        but1234_02.Checked = False
        but1234_03.Checked = False
        but1234_04.Checked = False
        but1234_05.Checked = False
        but1234_06.Checked = False
        but4250_01.Checked = False
        but4250_02.Checked = False
        but4250_03.Checked = False
        but4250_04.Checked = False
        but4250_05.Checked = False
        but4250_06.Checked = False
        but14123_01.Checked = False
        but14123_02.Checked = False
        but14123_03.Checked = False
        but14123_04.Checked = False
        but14123_05.Checked = False
        but14123_06.Checked = False
        but21094_01.Checked = False
        but21094_02.Checked = False
        but21094_03.Checked = False
        but21094_04.Checked = False
        but21094_05.Checked = False
        but21094_06.Checked = False
        but10279_01.Checked = False
        but10279_02.Checked = False
        but10279_03.Checked = False
        but10279_04.Checked = False
        but10279_05.Checked = False
        but10279_06.Checked = False
        but10910_01.Checked = False
        but10910_02.Checked = False
        but10910_03.Checked = False
        but10910_04.Checked = False
        but10910_05.Checked = False
        but10910_06.Checked = False
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB tirarports: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub deslabelar()
        On Error GoTo MOSTRARERRO
        labelatribuido.Font = New Font(Me.labelatribuido.Font, FontStyle.Regular)
        labelatribuido.Text = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB tirarports: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub organismus(ByVal umdeles As Short)
        On Error GoTo MOSTRARERRO
        Select Case umdeles
            Case 1, 46, 41, 42, 67, 23
                tirarports()
            Case 49, 45, 59
                but1474_01.Checked = True
                but1474_02.Checked = True
                but1474_03.Checked = True
                but1474_04.Checked = True
                but1474_05.Checked = True
                but1474_06.Checked = True
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("SUB tirarports: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

End Class





