Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb

Public Class EC
    Inherits Form

    Dim textolabelPort1RG As String
    Dim textolabelPort1RE As String
    Dim textolabelPort1RG_ As String
    Dim textolabelPort1RE_ As String
    Dim textolabelPort2RG As String
    Dim textolabelPort2RE As String
    Dim textolabelPort2RG_ As String
    Dim textolabelPort2RE_ As String

    Dim codigorow As basededadosDataSet.infarmedRow
    Dim codigoarray As New ArrayList
    Dim codigo4 As New meds

    Dim foco4 As String
    Dim novado As Boolean
    Dim mostraPVP As Double
    Dim mostraPR As Double
    Dim mostracomp As Double
    Dim mostracomp10 As Double
    Dim mostracomp15 As Double
    Dim mostracomp100R As Double
    Dim mostracompPort1RG As Double
    Dim mostracompPort1RE As Double
    Dim mostracompPort2RG As Double
    Dim mostracompPort2RE As Double
    Dim mostranome As String
    Dim mostradif15_10 As Double
    Dim mostradif100R_10 As Double
    Dim mostradifPort1RG_10 As Double
    Dim mostradifPort1RE_10 As Double
    Dim mostradifPort2RG_10 As Double
    Dim mostradifPort2RE_10 As Double
    Dim mostradif100R_15 As Double
    Dim mostradifPort1RG_15 As Double
    Dim mostradifPort1RE_15 As Double
    Dim mostradifPort2RG_15 As Double
    Dim mostradifPort2RE_15 As Double
    Dim mostradifPort1RG_100R As Double
    Dim mostradifPort1RE_100R As Double
    Dim mostradifPort2RG_100R As Double
    Dim mostradifPort2RE_100R As Double
    Dim mostradifPort1RE_Port1RG As Double
    Dim mostradifPort2RG_Port1RG As Double
    Dim mostradifPort2RE_Port1RG As Double
    Dim mostradifPort2RG_Port1RE As Double
    Dim mostradifPort2RE_Port1RE As Double
    Dim mostradifPort2RE_Port2RG As Double
    Dim intermedcodEC As Integer
    Dim intermedpvpEC As Double
    Dim intermedprEC As Double

    Dim qualport As String
    Dim qualport2 As String
    Dim portSN As Boolean
    Dim port2 As Boolean
    Dim portcomp As Double
    Dim portcomp2 As Double

    Dim infarmedTA4 As New basededadosDataSetTableAdapters.infarmedTableAdapter
    Dim DS4 As New basededadosDataSet
    Dim row4 As DataRow

    'esta variavel guarda o número de linhas lido da base de dados
    Dim infarmedTA4s As Integer = infarmedTA4.Fill(DS4.infarmed)


    Private Sub EC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo MOSTRARERRO
        Me.KeyPreview = True

        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub irbuscar4()
        On Error GoTo MOSTRARERRO
        
        codigorow = DS4.infarmed.FindBycode(codigo4.codigo)
        If Not IsNothing(codigorow) Then
            codigoarray.Add(codigorow)
            labelmed.Font = New Font(Me.labelmed.Font, FontStyle.Bold)
            intermedpvpEC = Replace(codigorow(17), ".", ",")
            intermedprEC = Replace(codigorow(18), ".", ",")
            pvpEC.Text = intermedpvpEC
            prEC.Text = intermedprEC
        Else
            aviadonexist()
        End If
        
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'valida se só contém algarismos. se não dá beep e fica onde está. se sim verifica se tem 7 algarismos e passa à frente. o que acontece ao fazer enter
    Private Sub frmDesigner_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        On Error GoTo MOSTRARERRO
        foco4 = Me.ActiveControl.Name()
        Dim soalgarismos As Boolean
        If Not Char.IsNumber(e.KeyChar) Then
            soalgarismos = False
        Else
            soalgarismos = True
        End If

        'Select Case soalgarismos
        '   Case False
        'Beep()
        'End Select


        If Asc(e.KeyChar) = Keys.Enter Then
            Select Case foco4
                Case "codEC"
                    If codEC.Text = "" Then
                        Beep()
                        My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
                    Else
                        Me.codEC.Focus()
                        codEC.SelectionStart = 0
                        codEC.SelectionLength = Len(codEC.Text)
                        limpar4()
                        codigo4.codigo = codEC.Text
                        incorporar()

                        mostrar()
                    End If
                Case "pvpEC"
                    novo(1)
                    mostraPVP = Replace(pvpEC.Text, ".", ",")
                    mostraPR = Replace(prEC.Text, ".", ",")
                    incorporar()
                    mostrar()
                Case "prEC"
                    novo(2)
                    mostraPR = Replace(prEC.Text, ".", ",")
                    mostraPVP = Replace(pvpEC.Text, ".", ",")
                    incorporar()
                    mostrar()
                Case Else
                    Me.codEC.Focus()
                    codEC.SelectionStart = 0
                    codEC.SelectionLength = Len(codEC.Text)
            End Select
        End If


        If Asc(e.KeyChar) = Keys.C Then
            My.Computer.Keyboard.SendKeys("{bs}")
            My.Computer.Keyboard.SendKeys("{ENTER}")
        End If



        If Asc(e.KeyChar) = Keys.Space Then

            limpar4()
            Me.codEC.Focus()

            codEC.SelectionStart = 0
            codEC.SelectionLength = Len(codEC.Text)

        End If



        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    'o próximo é o validador que está a funcionar. faz saltar o focus quando se inserem 7 caracteres. e lançam os calculos e amostragem
    Private Sub codEC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles codEC.TextChanged
        On Error GoTo MOSTRARERRO
        Dim Caracta44 As String
        Caracta44 = codEC.Text
        If Len(codEC.Text) = 7 Then
            If Caracta44 Like "#######" Then
                limpar4()
                codigo4.codigo = codEC.Text
                incorporar()
                mostrar()
                Me.codEC.Focus()
                codEC.SelectionStart = 0
                codEC.SelectionLength = Len(codEC.Text)
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub limpar4()
        On Error GoTo MOSTRARERRO
        'intermedpvpEC = 0
        'intermedprEC = 0
        novado = False
        tirarvalorcomp()
        tirarvalordif()
        tirardadosport()
        desboldar()
        mostraPR = 0
        mostraPVP = 0
        mostranome = ""
        labelmed.Text = ""
        textolabelPort1RG = ""
        textolabelPort1RE = ""
        textolabelPort1RG_ = ""
        textolabelPort1RE_ = ""
        textolabelPort2RG = ""
        textolabelPort2RE = ""
        textolabelPort2RG_ = ""
        textolabelPort2RE_ = ""
        LabelGEN.Text = ""
        labelPort1RG.Text = ""
        labelPort1RG_.Text = ""
        labelPort1RE.Text = ""
        labelPort1RE_.Text = ""
        labelPort2RG.Text = ""
        labelPort2RG_.Text = ""
        labelPort2RE.Text = ""
        comp10ec.Text = ""
        comp15ec.Text = ""
        comp100rec.Text = ""
        compPort1RG.Text = ""
        compPort1RE.Text = ""
        compPort2RG.Text = ""
        compPort2RE.Text = ""
        dif15_10.Text = ""
        dif100R_10.Text = ""
        difPort1RG_10.Text = ""
        difPort1RE_10.Text = ""
        difPort2RG_10.Text = ""
        difPort2RE_10.Text = ""
        dif100R_15.Text = ""
        difPort1RG_15.Text = ""
        difPort1RE_15.Text = ""
        difPort2RG_15.Text = ""
        difPort2RE_15.Text = ""
        difPort1RG_100R.Text = ""
        difPort1RE_100R.Text = ""
        difPort2RG_100R.Text = ""
        difPort2RE_100R.Text = ""
        difPort1RE_Port1RG.Text = ""
        difPort2RG_Port1RG.Text = ""
        difPort2RE_Port1RG.Text = ""
        difPort2RG_Port1RE.Text = ""
        difPort2RE_Port1RE.Text = ""
        difPort2RE_Port2RG.Text = ""
        pvpEC.Text = 0
        prEC.Text = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub mostrar()

        On Error GoTo MOSTRARERRO
        labelmed.Text = mostranome
        If Not IsNothing(codigorow) Then
            Select Case codigorow(7)
                Case True
                    LabelGEN.Text = "é GENÉRICO \"
                Case False
                    LabelGEN.Text = "NÃO é genérico /"
            End Select

            If mostracomp10 <= codigorow(17) Then
                comp10ec.Text = System.Math.Round(mostracomp10, 2)
            Else
                comp10ec.Text = codigorow(17)
            End If
            If mostracomp15 <= codigorow(17) Then
                comp15ec.Text = System.Math.Round(mostracomp15, 2)
            Else
                comp15ec.Text = codigorow(17)
            End If
            If mostracomp100R <= codigorow(17) Then
                comp100rec.Text = System.Math.Round(mostracomp100R, 2)
            Else
                comp100rec.Text = codigorow(17)
            End If
            If mostracompPort1RG <= codigorow(17) Then
                compPort1RG.Text = System.Math.Round(mostracompPort1RG, 2)
            Else
                compPort1RG.Text = codigorow(17)
            End If
            If mostracompPort1RE <= codigorow(17) Then
                compPort1RE.Text = System.Math.Round(mostracompPort1RE, 2)
            Else
                compPort1RE.Text = codigorow(17)
            End If
            If mostracompPort2RG <= codigorow(17) Then
                compPort2RG.Text = System.Math.Round(mostracompPort2RG, 2)
            Else
                compPort2RG.Text = codigorow(17)
            End If
            If mostracompPort2RE <= codigorow(17) Then
                compPort2RE.Text = System.Math.Round(mostracompPort2RE, 2)
            Else
                compPort2RE.Text = codigorow(17)
            End If
        Else
            comp10ec.Text = "0"
            comp15ec.Text = "0"
            comp100rec.Text = "0"
            compPort1RG.Text = "0"
            compPort1RE.Text = "0"
            compPort2RG.Text = "0"
            compPort2RE.Text = "0"
        End If
        dif15_10.Text = System.Math.Round(System.Math.Abs(mostradif15_10), 2)
        dif100R_10.Text = System.Math.Round(System.Math.Abs(mostradif100R_10), 2)
        difPort1RG_10.Text = System.Math.Round(System.Math.Abs(mostradifPort1RG_10), 2)
        difPort1RE_10.Text = System.Math.Round(System.Math.Abs(mostradifPort1RE_10), 2)
        difPort2RG_10.Text = System.Math.Round(System.Math.Abs(mostradifPort2RG_10), 2)
        difPort2RE_10.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_10), 2)

        dif100R_15.Text = System.Math.Round(System.Math.Abs(mostradif100R_15), 2)
        difPort1RG_15.Text = System.Math.Round(System.Math.Abs(mostradifPort1RG_15), 2)
        difPort1RE_15.Text = System.Math.Round(System.Math.Abs(mostradifPort1RE_15), 2)
        difPort2RG_15.Text = System.Math.Round(System.Math.Abs(mostradifPort2RG_15), 2)
        difPort2RE_15.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_15), 2)

        difPort1RG_100R.Text = System.Math.Round(System.Math.Abs(mostradifPort1RG_100R), 2)
        difPort1RE_100R.Text = System.Math.Round(System.Math.Abs(mostradifPort1RE_100R), 2)
        difPort2RG_100R.Text = System.Math.Round(System.Math.Abs(mostradifPort2RG_100R), 2)
        difPort2RE_100R.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_100R), 2)

        difPort1RE_Port1RG.Text = System.Math.Round(System.Math.Abs(mostradifPort1RE_Port1RG), 2)
        difPort2RG_Port1RG.Text = System.Math.Round(System.Math.Abs(mostradifPort2RG_Port1RG), 2)
        difPort2RE_Port1RG.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_Port1RG), 2)

        difPort2RG_Port1RE.Text = System.Math.Round(System.Math.Abs(mostradifPort2RG_Port1RE), 2)
        difPort2RE_Port1RE.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_Port1RE), 2)

        difPort2RE_Port2RG.Text = System.Math.Round(System.Math.Abs(mostradifPort2RE_Port2RG), 2)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub calc1015()
        On Error GoTo MOSTRARERRO
        If mostraPR > 0 And mostraPR < mostraPVP Then 'usa PR
            mostracomp10 = mostracomp * mostraPR

            If mostracomp <= 0.85 Then
                mostracomp15 = (mostracomp + 0.15) * mostraPR
            Else
                mostracomp15 = mostraPR
            End If
            mostracomp100R = mostraPR * 1.2

        Else 'usa PVP
            mostracomp10 = mostracomp * mostraPVP
            If mostracomp <= 0.85 Then
                mostracomp15 = (mostracomp + 0.15) * mostraPVP
            Else
                mostracomp15 = mostraPVP
            End If
            mostracomp100R = mostraPVP
        End If
        mostradif15_10 = mostracomp15 - mostracomp10
        mostradif100R_10 = mostracomp100R - mostracomp10
        mostradif100R_15 = mostracomp100R - mostracomp15
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub AtribCompPort(ByVal escolher As String, ByVal ordem As Short)
        On Error GoTo MOSTRARERRO
        Select Case ordem
            Case 1
                Select Case escolher
                    Case 4250
                        portcomp = 0.37
                    Case 1234
                        portcomp = 0.95
                    Case 10279
                        portcomp = 0.95
                    Case 10280
                        portcomp = 0.95
                    Case 10910
                        portcomp = 0.69
                    Case 14123
                        portcomp = 0.69
                    Case 147469
                        portcomp = 0.69
                    Case 1474100
                        portcomp = 1
                    Case 21094
                        portcomp = 1
                End Select
            Case 2
                Select Case escolher
                    Case 4250
                        portcomp2 = 0.37
                    Case 1234
                        portcomp2 = 0.95
                    Case 10279
                        portcomp2 = 0.95
                    Case 10280
                        portcomp2 = 0.95
                    Case 10910
                        portcomp2 = 0.69
                    Case 14123
                        portcomp2 = 0.69
                    Case 147469
                        portcomp2 = 0.69
                    Case 1474100
                        portcomp2 = 1
                    Case 21094
                        portcomp2 = 1
                End Select
        End Select
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub incorporar()
        On Error GoTo MOSTRARERRO
        If novado = "false" Then
            irbuscar4()
        End If

        If Not IsNothing(codigorow) Then
            labelmed.Font = New Font(Me.labelmed.Font, FontStyle.Bold)
            detectarport()
            mostranome = (codigorow(16)) & " (" & (codigorow(1)) & ") (" & (codigorow(3)) & ") (" & (codigorow(4)) & ") (" & (codigorow(2)) & ")"
            If novado = "false" Then

                mostraPVP = Replace(codigorow(17), ".", ",")
                mostraPR = Replace(codigorow(18), ".", ",")

            End If
            mostracomp = (codigorow(5) * 0.01)


            If portSN = False Then 'se não leva portaria
                portcomp = 0
                portcomp2 = 0

                calc1015()

            Else 'se leva portaria

                If qualport <> 147469 And qualport <> 1474100 Then 'se a port não é a 1474/2004
                    textolabelPort1RG = qualport & " RG"
                    textolabelPort1RE = qualport & " RE"
                    textolabelPort1RG_ = qualport & " RG"
                    textolabelPort1RE_ = qualport & " RE"
                    labelPort1RG.Text = textolabelPort1RG
                    labelPort1RE.Text = textolabelPort1RE
                    labelPort1RG_.Text = textolabelPort1RG_
                    labelPort1RE_.Text = textolabelPort1RE_
                Else 'se a port é a 1474/2004
                    If qualport = 147469 Then
                        textolabelPort1RG = "1474ad RG"
                        textolabelPort1RE = "1474ad RG"
                        textolabelPort1RG_ = "1474ad RG"
                        textolabelPort1RE_ = "1474ad RG"
                    ElseIf qualport = 1474100 Then
                        textolabelPort1RG = "1474nl RG"
                        textolabelPort1RE = "1474nl RG"
                        textolabelPort1RG_ = "1474nl RG"
                        textolabelPort1RE_ = "1474nl RG"
                    End If
                    labelPort1RG.Text = textolabelPort1RG
                    labelPort1RE.Text = textolabelPort1RE
                    labelPort1RG_.Text = textolabelPort1RG_
                    labelPort1RE_.Text = textolabelPort1RE_
                End If

                AtribCompPort(qualport, 1)

                calc1015()

                If mostraPR > 0 And mostraPR < mostraPVP Then 'usa PR
                    mostracompPort1RG = portcomp * mostraPR
                    If portcomp <= 0.85 Then
                        mostracompPort1RE = (portcomp + 0.15) * mostraPR
                    Else
                        mostracompPort1RE = mostraPR
                    End If


                Else 'usa PVP
                    mostracompPort1RG = portcomp * mostraPVP
                    If portcomp <= 0.85 Then
                        mostracompPort1RE = (portcomp + 0.15) * mostraPVP
                    Else
                        mostracompPort1RE = mostraPVP
                    End If
                End If

                If port2 = True Then 'se pode levar mais do que uma portaria

                    If qualport2 <> 147469 And qualport <> 1474100 Then 'se a port não é a 1474/2004
                        textolabelPort2RG = qualport2 & " RG"
                        textolabelPort2RE = qualport2 & " RE"
                        textolabelPort2RG_ = qualport2 & " RG"
                        labelPort2RG.Text = textolabelPort2RG
                        labelPort2RE.Text = textolabelPort2RE
                        labelPort2RG_.Text = textolabelPort2RG_
                    Else 'se a port é a 1474/2004
                        If qualport2 = 147469 Then
                            textolabelPort2RG = "1474ad RG"
                            textolabelPort2RE = "1474ad RG"
                            textolabelPort2RG_ = "1474ad RG"
                        ElseIf qualport2 = 1474100 Then
                            textolabelPort2RG = "1474nl RG"
                            textolabelPort2RE = "1474nl RG"
                            textolabelPort2RG_ = "1474nl RG"
                        End If
                        labelPort2RG.Text = textolabelPort2RG
                        labelPort2RE.Text = textolabelPort2RE
                        labelPort2RG_.Text = textolabelPort2RG_
                    End If


                    AtribCompPort(qualport2, 2)

                    calc1015()

                    If mostraPR > 0 And mostraPR < mostraPVP Then 'usa PR
                        mostracompPort2RG = portcomp2 * mostraPR

                        If portcomp2 <= 0.85 Then
                            mostracompPort2RE = (portcomp2 + 0.15) * mostraPR
                        Else
                            mostracompPort2RE = mostraPR
                        End If


                    Else 'usa PVP

                        mostracompPort2RG = portcomp2 * mostraPVP

                        If portcomp2 <= 0.85 Then
                            mostracompPort2RE = (portcomp2 + 0.15) * mostraPVP
                        Else
                            mostracompPort2RE = mostraPVP
                        End If

                    End If

                Else 'se só levar uma portaria
                    portcomp2 = mostracomp
                End If

                mostradifPort1RG_10 = mostracompPort1RG - mostracomp10
                mostradifPort1RE_10 = mostracompPort1RE - mostracomp10
                mostradifPort2RG_10 = mostracompPort2RG - mostracomp10
                mostradifPort2RE_10 = mostracompPort2RE - mostracomp10

                mostradifPort1RG_15 = mostracompPort1RG - mostracomp15
                mostradifPort1RE_15 = mostracompPort1RE - mostracomp15
                mostradifPort2RG_15 = mostracompPort2RG - mostracomp15
                mostradifPort2RE_15 = mostracompPort2RE - mostracomp15

                mostradifPort1RG_100R = mostracompPort1RG - mostracomp100R
                mostradifPort1RE_100R = mostracompPort1RE - mostracomp100R
                mostradifPort2RG_100R = mostracompPort2RG - mostracomp100R
                mostradifPort2RE_100R = mostracompPort2RE - mostracomp100R

                mostradifPort1RE_Port1RG = mostracompPort1RE - mostracompPort1RG
                mostradifPort2RG_Port1RG = mostracompPort2RG - mostracompPort1RG
                mostradifPort2RE_Port1RG = mostracompPort2RE - mostracompPort1RG

                mostradifPort2RG_Port1RE = mostracompPort2RG - mostracompPort1RE
                mostradifPort2RE_Port1RE = mostracompPort2RE - mostracompPort1RE

                mostradifPort2RE_Port2RG = mostracompPort2RE - mostracompPort2RG

            End If
        Else
            aviadonexist()
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub detectarport()
        On Error GoTo MOSTRARERRO
        If Not IsNothing(codigorow) Then
            If codigorow(9) = True Then

                If portSN = True Then
                    port2 = True
                    qualport2 = "4250"
                Else
                    portSN = True
                    qualport = "4250"
                End If
            End If

            If codigorow(10) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "1234"
                    Else
                        portSN = True
                        qualport = "1234"
                    End If
                End If
            End If

            If codigorow(11) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "10279"
                    Else
                        portSN = True
                        qualport = "10279"
                    End If
                End If
            End If

            If codigorow(12) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "10280"
                    Else
                        portSN = True
                        qualport = "10280"
                    End If
                End If
            End If

            If codigorow(13) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "10910"
                    Else
                        portSN = True
                        qualport = "10910"
                    End If
                End If
            End If

            If codigorow(14) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "14123"
                    Else
                        portSN = True
                        qualport = "14123"
                    End If
                End If
            End If

            If codigorow(15) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "147469"
                    Else
                        portSN = True
                        qualport = "147469"
                    End If
                End If
            End If

            If codigorow(19) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "21094"
                    Else
                        portSN = True
                        qualport = "21094"
                    End If
                End If
            End If

            If codigorow(20) = True Then

                If codigorow(5).ToString <> 0 Then
                    If portSN = True Then
                        port2 = True
                        qualport2 = "1474100"
                    Else
                        portSN = True
                        qualport = "1474100"
                    End If
                End If
            End If
        End If
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub dif15_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dif15_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        dif15_10.ForeColor = Color.Red
        dif15_10.Font = New Font(Me.dif15_10.Font, FontStyle.Bold)
        labelRE.ForeColor = Color.Red
        labelRE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub dif100R_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dif100R_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        dif100R_10.ForeColor = Color.Red
        dif100R_10.Font = New Font(Me.dif100R_10.Font, FontStyle.Bold)
        label100R.ForeColor = Color.Red
        label100R.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difport1RG_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RG_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RG_10.ForeColor = Color.Red
        difPort1RG_10.Font = New Font(Me.difPort1RG_10.Font, FontStyle.Bold)
        labelPort1RG.ForeColor = Color.Red
        labelPort1RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difport1RE_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RE_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RE_10.ForeColor = Color.Red
        difPort1RE_10.Font = New Font(Me.difPort1RE_10.Font, FontStyle.Bold)
        labelPort1RE.ForeColor = Color.Red
        labelPort1RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difport2RG_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RG_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RG_10.ForeColor = Color.Red
        difPort2RG_10.Font = New Font(Me.difPort2RG_10.Font, FontStyle.Bold)
        labelPort2RG.ForeColor = Color.Red
        labelPort2RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difport2RE_10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_10.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_10.ForeColor = Color.Red
        difPort2RE_10.Font = New Font(Me.difPort2RE_10.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRG_.ForeColor = Color.Red
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub dif100R_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dif100R_15.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        dif100R_15.ForeColor = Color.Red
        dif100R_15.Font = New Font(Me.dif100R_15.Font, FontStyle.Bold)
        label100R.ForeColor = Color.Red
        label100R.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRE_.ForeColor = Color.Red
        labelRE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort1RG_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RG_15.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RG_15.ForeColor = Color.Red
        difPort1RG_15.Font = New Font(Me.difPort1RG_15.Font, FontStyle.Bold)
        labelPort1RG.ForeColor = Color.Red
        labelPort1RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRE_.ForeColor = Color.Red
        labelRE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort1RE_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RE_15.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RE_15.ForeColor = Color.Red
        difPort1RE_15.Font = New Font(Me.difPort1RE_15.Font, FontStyle.Bold)
        labelPort1RE.ForeColor = Color.Red
        labelPort1RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRE_.ForeColor = Color.Red
        labelRE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2RG_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RG_15.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RG_15.ForeColor = Color.Red
        difPort2RG_15.Font = New Font(Me.difPort2RG_15.Font, FontStyle.Bold)
        labelPort2RG.ForeColor = Color.Red
        labelPort2RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRE_.ForeColor = Color.Red
        labelRE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Re_15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_15.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_15.ForeColor = Color.Red
        difPort2RE_15.Font = New Font(Me.difPort2RE_15.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelRE_.ForeColor = Color.Red
        labelRE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort1RG_100R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RG_100R.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RG_100R.ForeColor = Color.Red
        difPort1RG_100R.Font = New Font(Me.difPort1RG_100R.Font, FontStyle.Bold)
        labelPort1RG.ForeColor = Color.Red
        labelPort1RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        label100R_.ForeColor = Color.Red
        label100R_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort1Re_100R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RE_100R.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RE_100R.ForeColor = Color.Red
        difPort1RE_100R.Font = New Font(Me.difPort1RE_100R.Font, FontStyle.Bold)
        labelPort1RE.ForeColor = Color.Red
        labelPort1RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        label100R_.ForeColor = Color.Red
        label100R_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Rg_100R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RG_100R.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RG_100R.ForeColor = Color.Red
        difPort2RG_100R.Font = New Font(Me.difPort2RG_100R.Font, FontStyle.Bold)
        labelPort2RG.ForeColor = Color.Red
        labelPort2RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        label100R_.ForeColor = Color.Red
        label100R_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Re_100R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_100R.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_100R.ForeColor = Color.Red
        difPort2RE_100R.Font = New Font(Me.difPort2RE_100R.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        label100R_.ForeColor = Color.Red
        label100R_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort1Re_port1rg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort1RE_Port1RG.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort1RE_Port1RG.ForeColor = Color.Red
        difPort1RE_Port1RG.Font = New Font(Me.difPort1RE_Port1RG.Font, FontStyle.Bold)
        labelPort1RE.ForeColor = Color.Red
        labelPort1RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort1RG_.ForeColor = Color.Red
        labelPort1RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Rg_port1rg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RG_Port1RG.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RG_Port1RG.ForeColor = Color.Red
        difPort2RG_Port1RG.Font = New Font(Me.difPort2RG_Port1RG.Font, FontStyle.Bold)
        labelPort2RG.ForeColor = Color.Red
        labelPort2RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort1RG_.ForeColor = Color.Red
        labelPort1RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Re_port1rg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_Port1RG.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_Port1RG.ForeColor = Color.Red
        difPort2RE_Port1RG.Font = New Font(Me.difPort2RE_Port1RG.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort1RG_.ForeColor = Color.Red
        labelPort1RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Rg_port1re_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RG_Port1RE.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RG_Port1RE.ForeColor = Color.Red
        difPort2RG_Port1RE.Font = New Font(Me.difPort2RG_Port1RE.Font, FontStyle.Bold)
        labelPort2RG.ForeColor = Color.Red
        labelPort2RG.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort1RE_.ForeColor = Color.Red
        labelPort1RE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Re_port1re_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_Port1RE.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_Port1RE.ForeColor = Color.Red
        difPort2RE_Port1RE.Font = New Font(Me.difPort2RE_Port1RE.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort1RE_.ForeColor = Color.Red
        labelPort1RE_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Private Sub difPort2Re_port2rg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles difPort2RE_Port2RG.Click
        On Error GoTo MOSTRARERRO
        desboldar()
        difPort2RE_Port2RG.ForeColor = Color.Red
        difPort2RE_Port2RG.Font = New Font(Me.difPort2RE_Port2RG.Font, FontStyle.Bold)
        labelPort2RE.ForeColor = Color.Red
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Bold)
        labelPort2RG_.ForeColor = Color.Red
        labelPort2RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Bold)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub desboldar()
        On Error GoTo MOSTRARERRO
        codEC.BackColor = Color.White
        labelmed.ForeColor = Color.Black
        labelmed.Font = New Font(Me.labelmed.Font, FontStyle.Regular)
        dif15_10.ForeColor = Color.Black
        dif15_10.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        labelRE.ForeColor = Color.Black
        labelRE.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        labelRG_.ForeColor = Color.Black
        labelRG_.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        labelRG.ForeColor = Color.Black
        labelRG.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        labelRE_.ForeColor = Color.Black
        labelRE_.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        label100R_.ForeColor = Color.Black
        label100R_.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        label100R.ForeColor = Color.Black
        label100R.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        labelPort1RG.ForeColor = Color.Black
        labelPort1RG.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        labelPort1RG_.ForeColor = Color.Black
        labelPort1RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        labelPort1RE.ForeColor = Color.Black
        labelPort1RE.Font = New Font(Me.labelPort1RE.Font, FontStyle.Regular)
        labelPort1RE_.ForeColor = Color.Black
        labelPort1RE_.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        labelPort2RG_.ForeColor = Color.Black
        labelPort2RG_.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        labelPort2RG.ForeColor = Color.Black
        labelPort2RG.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        labelPort2RE.ForeColor = Color.Black
        labelPort2RE.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        dif100R_10.ForeColor = Color.Black
        dif100R_10.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort1RG_10.ForeColor = Color.Black
        difPort1RG_10.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        difPort1RE_10.ForeColor = Color.Black
        difPort1RE_10.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort2RG_10.ForeColor = Color.Black
        difPort2RG_10.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort2RE_10.ForeColor = Color.Black
        difPort2RE_10.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        dif100R_15.ForeColor = Color.Black
        dif100R_15.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort1RG_15.ForeColor = Color.Black
        difPort1RG_15.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort1RE_15.ForeColor = Color.Black
        difPort1RE_15.Font = New Font(Me.dif15_10.Font, FontStyle.Regular)
        difPort2RG_15.ForeColor = Color.Black
        difPort2RG_15.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort2RE_15.ForeColor = Color.Black
        difPort2RE_15.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort1RG_100R.ForeColor = Color.Black
        difPort1RG_100R.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort1RE_100R.ForeColor = Color.Black
        difPort1RE_100R.Font = New Font(Me.dif100R_10.Font, FontStyle.Regular)
        difPort2RG_100R.ForeColor = Color.Black
        difPort2RG_100R.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort2RE_100R.ForeColor = Color.Black
        difPort2RE_100R.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort1RE_Port1RG.ForeColor = Color.Black
        difPort1RE_Port1RG.Font = New Font(Me.difPort1RE_10.Font, FontStyle.Regular)
        difPort2RG_Port1RG.ForeColor = Color.Black
        difPort2RG_Port1RG.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort2RE_Port1RG.ForeColor = Color.Black
        difPort2RE_Port1RG.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort2RG_Port1RE.ForeColor = Color.Black
        difPort2RG_Port1RE.Font = New Font(Me.labelRE.Font, FontStyle.Regular)
        difPort2RE_Port1RE.ForeColor = Color.Black
        difPort2RE_Port1RE.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        difPort2RE_Port2RG.ForeColor = Color.Black
        difPort2RE_Port2RG.Font = New Font(Me.labelRG_.Font, FontStyle.Regular)
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub tirarvalorcomp()
        On Error GoTo MOSTRARERRO
        mostracomp = 0
        mostracomp10 = 0
        mostracomp100R = 0
        mostracomp15 = 0
        mostracompPort1RE = 0
        mostracompPort1RG = 0
        mostracompPort2RG = 0
        mostracompPort2RE = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub tirarvalordif()
        On Error GoTo MOSTRARERRO
        mostradif100R_10 = 0
        mostradif100R_15 = 0
        mostradif15_10 = 0

        mostradifPort1RG_10 = 0
        mostradifPort1RE_10 = 0
        mostradifPort2RG_10 = 0
        mostradifPort2RE_10 = 0

        mostradifPort1RG_15 = 0
        mostradifPort1RE_15 = 0
        mostradifPort2RG_15 = 0
        mostradifPort2RE_15 = 0

        mostradifPort1RG_100R = 0
        mostradifPort1RE_100R = 0
        mostradifPort2RG_100R = 0
        mostradifPort2RE_100R = 0

        mostradifPort1RE_Port1RG = 0
        mostradifPort2RG_Port1RG = 0
        mostradifPort2RE_Port1RG = 0

        mostradifPort2RG_Port1RE = 0
        mostradifPort2RE_Port1RE = 0

        mostradifPort2RE_Port2RG = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub tirardadosport()
        On Error GoTo MOSTRARERRO
        portSN = False
        port2 = False
        qualport = 0
        qualport2 = 0
        portcomp = 0
        portcomp2 = 0
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub aviadonexist()
        On Error GoTo MOSTRARERRO
        codEC.BackColor = Color.Red
        labelmed.ForeColor = Color.Red
        labelmed.Font = New Font(Me.labelmed.Font, FontStyle.Bold)
        mostranome = "o código aviado não existe"
        labelmed.Text = mostranome
        intermedpvpEC = 0
        intermedprEC = 0
        pvpEC.Text = ""
        prEC.Text = ""
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub

    Sub novo(ByVal caixa As Short)
        On Error GoTo MOSTRARERRO
        Select Case caixa
            Case 2
                intermedprEC = prEC.Text
            Case 1
                intermedpvpEC = pvpEC.Text
        End Select
        novado = True
        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub
End Class