Public Class bd

    Private Sub bd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'BasededadosDataSet1.infarmed' table. You can move, or remove it, as needed.
        Me.InfarmedTableAdapter1.Fill(Me.BasededadosDataSet1.infarmed)
        'TODO: This line of code loads data into the 'BasededadosDataSet.infarmed' table. You can move, or remove it, as needed.
        'Me.InfarmedTableAdapter.Fill(Me.BasededadosDataSet.infarmed)

    End Sub


    Private Sub FillToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.InfarmedTableAdapter1.Fill(Me.BasededadosDataSet1.infarmed)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub CódigoToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CódigoToolStripButton.Click
        Try
            Me.InfarmedTableAdapter.código(Me.BasededadosDataSet.infarmed, New System.Nullable(Of Decimal)(CType(CodeToolStripTextBox.Text, Decimal)))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub nomeToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nomeToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.nome(Me.BasededadosDataSet1.infarmed, nomeToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub DciToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DciToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.dci(Me.BasededadosDataSet1.infarmed, DciToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub GHToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GHToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.GH(Me.BasededadosDataSet1.infarmed, GhToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub CompToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.comp(Me.BasededadosDataSet1.infarmed, New System.Nullable(Of Decimal)(CType(CompToolStripTextBox.Text, Decimal)))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub



    Private Sub DosagemToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DosagemToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.dosagem(Me.BasededadosDataSet1.infarmed, DoseToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub



    'Private Sub ComparticipaçãoToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   Try
    '      Me.InfarmedTableAdapter.comparticipação(Me.BasededadosDataSet.infarmed, New System.Nullable(Of Decimal)(CType(CompToolStripTextBox.Text, Decimal)))
    ' Catch ex As System.Exception
    '    System.Windows.Forms.MessageBox.Show(ex.Message)
    ' End Try
    '
    '    End Sub

    Private Sub QuantidadeToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuantidadeToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.quantidade(Me.BasededadosDataSet1.infarmed, QtyToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub LaboratórioToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaboratórioToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.Laboratório(Me.BasededadosDataSet1.infarmed, LabToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub




    'Private Sub FillBy1ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   Try
    '      Me.InfarmedTableAdapter.FillBy1(Me.BasededadosDataSet.infarmed, CType(Desp_10279ToolStripTextBox.Text, Boolean))
    ' Catch ex As System.Exception
    '    System.Windows.Forms.MessageBox.Show(ex.Message)
    'End Try
    '
    '   End Sub

    Private Sub _10279ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _10279ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter._10279(Me.BasededadosDataSet.infarmed, CType(Desp_10279ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub _10280ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _10280ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter._10280(Me.BasededadosDataSet.infarmed, CType(Desp_10280ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub






    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   Try
    '      Me.InfarmedTableAdapter.dci(Me.BasededadosDataSet.infarmed, DciToolStripTextBox.Text)
    ' Catch ex As System.Exception
    '    System.Windows.Forms.MessageBox.Show(ex.Message)
    ' End Try
    'End Sub

    Private Sub FormaToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormaToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.forma(Me.BasededadosDataSet1.infarmed, FormaToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub



    Private Sub GenéricoToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenéricoToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.genérico(Me.BasededadosDataSet1.infarmed, CType(GenToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub








    Private Sub PVPToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PVPToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.PVP(Me.BasededadosDataSet1.infarmed, PvpToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub PRToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PRToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.PR(Me.BasededadosDataSet1.infarmed, PrToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub



    'Private Sub PVUToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PVUToolStripButton.Click
    '   Try
    '      Dim Param1 As Object = New Object
    '     Me.InfarmedTableAdapter.PVU(Me.BasededadosDataSet.infarmed, Param1)
    'Catch ex As System.Exception
    '   System.Windows.Forms.MessageBox.Show(ex.Message)
    'End Try
    '
    'End Sub

    ' Private Sub NL1474ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NL1474ToolStripButton.Click
    '    Try
    '        Me.InfarmedTableAdapter1.NL1474(Me.BasededadosDataSet1.infarmed, CType(NL1474ToolStripTextBox.Text, Boolean))
    '   Catch ex As System.Exception
    '       System.Windows.Forms.MessageBox.Show(ex.Message)
    '   End Try
    '
    '    End Sub





    ' Private Sub P1474ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles P1474ToolStripButton.Click
    '    Try
    '       Me.InfarmedTableAdapter1.p1474(Me.BasededadosDataSet1.infarmed, CType(Port_1474ToolStripTextBox.Text, Boolean))
    '  Catch ex As System.Exception
    '     System.Windows.Forms.MessageBox.Show(ex.Message)
    'End Try

    '    End Sub



    Private Sub D4250ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles D4250ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.d4250(Me.BasededadosDataSet1.infarmed, CType(Desp_4250ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub D1234ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles D1234ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.d1234(Me.BasededadosDataSet1.infarmed, CType(Desp_1234ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub D21094ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles d21094ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.d21094(Me.BasededadosDataSet1.infarmed, CType(Desp_21094ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub D14123ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles D14123ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.D14123(Me.BasededadosDataSet1.infarmed, CType(Desp_14123ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub D10910ToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles D10910ToolStripButton.Click
        Try
            Me.InfarmedTableAdapter1.D10910(Me.BasededadosDataSet1.infarmed, CType(Desp_10910ToolStripTextBox.Text, Boolean))
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub





    Private Sub pesquisar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'Me.InfarmedTableAdapter.d10910(Me.BasededadosDataSet.infarmed, Desp_10910ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.d14123(Me.BasededadosDataSet.infarmed, Desp_14123ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.d21249(Me.BasededadosDataSet.infarmed, Desp_21249ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.d1234(Me.BasededadosDataSet.infarmed, Desp_1234ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.d4250(Me.BasededadosDataSet.infarmed, Desp_4250ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.p1474(Me.BasededadosDataSet.infarmed, Port_1474ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.código(Me.BasededadosDataSet.infarmed, CodeToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.dci(Me.BasededadosDataSet.infarmed, DciToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.forma(Me.BasededadosDataSet.infarmed, FormaToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.GH(Me.BasededadosDataSet.infarmed, GhToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.comp(Me.BasededadosDataSet.infarmed, CompToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.genérico(Me.BasededadosDataSet.infarmed, GenToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.Laboratório(Me.BasededadosDataSet.infarmed, LabToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.Quantidade(Me.BasededadosDataSet.infarmed, QtyToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.PR(Me.BasededadosDataSet.infarmed, PrToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.PVP(Me.BasededadosDataSet.infarmed, PvpToolStripTextBox.Text)
            'Me.InfarmedTableAdapter._10279(Me.BasededadosDataSet.infarmed, Desp_10279ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter._10280(Me.BasededadosDataSet.infarmed, Desp_10280ToolStripTextBox.Text)
            'Me.InfarmedTableAdapter.dosagem(Me.BasededadosDataSet.infarmed, DoseToolStripTextBox.Text)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub



    Private Sub limpar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Desp_10910ToolStripTextBox.Text = ""
        Desp_14123ToolStripTextBox.Text = ""
        Desp_21094ToolStripTextBox.Text = ""
        Desp_1234ToolStripTextBox.Text = ""
        Desp_4250ToolStripTextBox.Text = ""
        'AQUIPort_1474ToolStripTextBox.Text = ""
        CodeToolStripTextBox.Text = ""
        DciToolStripTextBox.Text = ""
        FormaToolStripTextBox.Text = ""
        GhToolStripTextBox.Text = ""
        CompToolStripTextBox.Text = ""
        DoseToolStripTextBox.Text = ""
        LabToolStripTextBox.Text = ""
        QtyToolStripTextBox.Text = ""
        GenToolStripTextBox.Text = ""
        PvpToolStripTextBox.Text = ""
        PrToolStripTextBox.Text = ""
        Desp_10279ToolStripTextBox.Text = ""
        Desp_10280ToolStripTextBox.Text = ""
    End Sub




End Class