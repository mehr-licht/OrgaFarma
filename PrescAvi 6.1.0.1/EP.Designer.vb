<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EP
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.labelPort1RE = New System.Windows.Forms.Label
        Me.labelPort1RG = New System.Windows.Forms.Label
        Me.label100R = New System.Windows.Forms.Label
        Me.labelRG = New System.Windows.Forms.Label
        Me.labelRE = New System.Windows.Forms.Label
        Me.labelPVP2 = New System.Windows.Forms.Label
        Me.labelPVP1 = New System.Windows.Forms.Label
        Me.PVP2 = New System.Windows.Forms.MaskedTextBox
        Me.pvp1 = New System.Windows.Forms.MaskedTextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.difPVP = New System.Windows.Forms.MaskedTextBox
        Me.PRdoPVP = New System.Windows.Forms.MaskedTextBox
        Me.labelPR = New System.Windows.Forms.Label
        Me.labeldifPVP = New System.Windows.Forms.Label
        Me.labeldifPR = New System.Windows.Forms.Label
        Me.labelPVPdoPR = New System.Windows.Forms.Label
        Me.PVPdoPR = New System.Windows.Forms.MaskedTextBox
        Me.difPR = New System.Windows.Forms.MaskedTextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.labelPR2 = New System.Windows.Forms.Label
        Me.labelPR1 = New System.Windows.Forms.Label
        Me.PR2 = New System.Windows.Forms.MaskedTextBox
        Me.PR1 = New System.Windows.Forms.MaskedTextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.codigodoPVP = New System.Windows.Forms.TextBox
        Me.codigodoPR = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'labelPort1RE
        '
        Me.labelPort1RE.AutoSize = True
        Me.labelPort1RE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPort1RE.Location = New System.Drawing.Point(44, 455)
        Me.labelPort1RE.Name = "labelPort1RE"
        Me.labelPort1RE.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labelPort1RE.Size = New System.Drawing.Size(0, 13)
        Me.labelPort1RE.TabIndex = 153
        '
        'labelPort1RG
        '
        Me.labelPort1RG.AutoSize = True
        Me.labelPort1RG.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPort1RG.Location = New System.Drawing.Point(44, 510)
        Me.labelPort1RG.Name = "labelPort1RG"
        Me.labelPort1RG.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labelPort1RG.Size = New System.Drawing.Size(0, 13)
        Me.labelPort1RG.TabIndex = 152
        '
        'label100R
        '
        Me.label100R.AutoSize = True
        Me.label100R.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label100R.Location = New System.Drawing.Point(44, 379)
        Me.label100R.Name = "label100R"
        Me.label100R.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.label100R.Size = New System.Drawing.Size(113, 13)
        Me.label100R.TabIndex = 151
        Me.label100R.Text = "Genérico 100% (R)"
        '
        'labelRG
        '
        Me.labelRG.AutoSize = True
        Me.labelRG.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelRG.Location = New System.Drawing.Point(44, 259)
        Me.labelRG.Name = "labelRG"
        Me.labelRG.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labelRG.Size = New System.Drawing.Size(83, 13)
        Me.labelRG.TabIndex = 150
        Me.labelRG.Text = "Regime Geral"
        '
        'labelRE
        '
        Me.labelRE.AutoSize = True
        Me.labelRE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelRE.Location = New System.Drawing.Point(44, 319)
        Me.labelRE.Name = "labelRE"
        Me.labelRE.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labelRE.Size = New System.Drawing.Size(101, 13)
        Me.labelRE.TabIndex = 149
        Me.labelRE.Text = "Regime Especial"
        '
        'labelPVP2
        '
        Me.labelPVP2.AutoSize = True
        Me.labelPVP2.Location = New System.Drawing.Point(196, 51)
        Me.labelPVP2.Name = "labelPVP2"
        Me.labelPVP2.Size = New System.Drawing.Size(34, 13)
        Me.labelPVP2.TabIndex = 148
        Me.labelPVP2.Text = "PVP2"
        '
        'labelPVP1
        '
        Me.labelPVP1.AutoSize = True
        Me.labelPVP1.Location = New System.Drawing.Point(93, 51)
        Me.labelPVP1.Name = "labelPVP1"
        Me.labelPVP1.Size = New System.Drawing.Size(34, 13)
        Me.labelPVP1.TabIndex = 147
        Me.labelPVP1.Text = "PVP1"
        '
        'PVP2
        '
        Me.PVP2.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PVP2.Location = New System.Drawing.Point(173, 72)
        Me.PVP2.Name = "PVP2"
        Me.PVP2.Size = New System.Drawing.Size(77, 32)
        Me.PVP2.TabIndex = 146
        Me.PVP2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pvp1
        '
        Me.pvp1.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pvp1.Location = New System.Drawing.Point(68, 73)
        Me.pvp1.Name = "pvp1"
        Me.pvp1.Size = New System.Drawing.Size(77, 32)
        Me.pvp1.TabIndex = 145
        Me.pvp1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(44, 629)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(0, 13)
        Me.Label3.TabIndex = 155
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(44, 569)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(0, 13)
        Me.Label4.TabIndex = 154
        '
        'difPVP
        '
        Me.difPVP.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.difPVP.Location = New System.Drawing.Point(122, 111)
        Me.difPVP.Name = "difPVP"
        Me.difPVP.Size = New System.Drawing.Size(77, 32)
        Me.difPVP.TabIndex = 156
        Me.difPVP.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PRdoPVP
        '
        Me.PRdoPVP.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PRdoPVP.Location = New System.Drawing.Point(122, 208)
        Me.PRdoPVP.Name = "PRdoPVP"
        Me.PRdoPVP.Size = New System.Drawing.Size(77, 32)
        Me.PRdoPVP.TabIndex = 157
        Me.PRdoPVP.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'labelPR
        '
        Me.labelPR.AutoSize = True
        Me.labelPR.Location = New System.Drawing.Point(145, 192)
        Me.labelPR.Name = "labelPR"
        Me.labelPR.Size = New System.Drawing.Size(22, 13)
        Me.labelPR.TabIndex = 158
        Me.labelPR.Text = "PR"
        '
        'labeldifPVP
        '
        Me.labeldifPVP.AutoSize = True
        Me.labeldifPVP.Location = New System.Drawing.Point(145, 146)
        Me.labeldifPVP.Name = "labeldifPVP"
        Me.labeldifPVP.Size = New System.Drawing.Size(42, 13)
        Me.labeldifPVP.TabIndex = 159
        Me.labeldifPVP.Text = "dif PVP"
        '
        'labeldifPR
        '
        Me.labeldifPR.AutoSize = True
        Me.labeldifPR.Location = New System.Drawing.Point(535, 146)
        Me.labeldifPR.Name = "labeldifPR"
        Me.labeldifPR.Size = New System.Drawing.Size(36, 13)
        Me.labeldifPR.TabIndex = 174
        Me.labeldifPR.Text = "dif PR"
        '
        'labelPVPdoPR
        '
        Me.labelPVPdoPR.AutoSize = True
        Me.labelPVPdoPR.Location = New System.Drawing.Point(535, 192)
        Me.labelPVPdoPR.Name = "labelPVPdoPR"
        Me.labelPVPdoPR.Size = New System.Drawing.Size(28, 13)
        Me.labelPVPdoPR.TabIndex = 173
        Me.labelPVPdoPR.Text = "PVP"
        '
        'PVPdoPR
        '
        Me.PVPdoPR.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PVPdoPR.Location = New System.Drawing.Point(512, 208)
        Me.PVPdoPR.Name = "PVPdoPR"
        Me.PVPdoPR.Size = New System.Drawing.Size(77, 32)
        Me.PVPdoPR.TabIndex = 172
        Me.PVPdoPR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'difPR
        '
        Me.difPR.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.difPR.Location = New System.Drawing.Point(512, 111)
        Me.difPR.Name = "difPR"
        Me.difPR.Size = New System.Drawing.Size(77, 32)
        Me.difPR.TabIndex = 171
        Me.difPR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(434, 629)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(0, 13)
        Me.Label8.TabIndex = 170
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(434, 569)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(0, 13)
        Me.Label9.TabIndex = 169
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(434, 455)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(0, 13)
        Me.Label10.TabIndex = 168
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(434, 510)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(0, 13)
        Me.Label11.TabIndex = 167
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(434, 379)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(113, 13)
        Me.Label12.TabIndex = 166
        Me.Label12.Text = "Genérico 100% (R)"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(434, 259)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(83, 13)
        Me.Label13.TabIndex = 165
        Me.Label13.Text = "Regime Geral"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(434, 319)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(101, 13)
        Me.Label14.TabIndex = 164
        Me.Label14.Text = "Regime Especial"
        '
        'labelPR2
        '
        Me.labelPR2.AutoSize = True
        Me.labelPR2.Location = New System.Drawing.Point(586, 51)
        Me.labelPR2.Name = "labelPR2"
        Me.labelPR2.Size = New System.Drawing.Size(28, 13)
        Me.labelPR2.TabIndex = 163
        Me.labelPR2.Text = "PR2"
        '
        'labelPR1
        '
        Me.labelPR1.AutoSize = True
        Me.labelPR1.Location = New System.Drawing.Point(483, 51)
        Me.labelPR1.Name = "labelPR1"
        Me.labelPR1.Size = New System.Drawing.Size(28, 13)
        Me.labelPR1.TabIndex = 162
        Me.labelPR1.Text = "PR1"
        '
        'PR2
        '
        Me.PR2.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PR2.Location = New System.Drawing.Point(563, 72)
        Me.PR2.Name = "PR2"
        Me.PR2.Size = New System.Drawing.Size(77, 32)
        Me.PR2.TabIndex = 161
        Me.PR2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PR1
        '
        Me.PR1.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PR1.Location = New System.Drawing.Point(458, 73)
        Me.PR1.Name = "PR1"
        Me.PR1.Size = New System.Drawing.Size(77, 32)
        Me.PR1.TabIndex = 160
        Me.PR1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.WindowText
        Me.GroupBox1.Location = New System.Drawing.Point(362, -2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(10, 682)
        Me.GroupBox1.TabIndex = 175
        Me.GroupBox1.TabStop = False
        '
        'codigodoPVP
        '
        Me.codigodoPVP.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.codigodoPVP.Location = New System.Drawing.Point(109, 11)
        Me.codigodoPVP.Name = "codigodoPVP"
        Me.codigodoPVP.Size = New System.Drawing.Size(100, 31)
        Me.codigodoPVP.TabIndex = 178
        '
        'codigodoPR
        '
        Me.codigodoPR.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.codigodoPR.Location = New System.Drawing.Point(501, 11)
        Me.codigodoPR.Name = "codigodoPR"
        Me.codigodoPR.Size = New System.Drawing.Size(100, 31)
        Me.codigodoPR.TabIndex = 179
        '
        'EP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(728, 671)
        Me.Controls.Add(Me.codigodoPR)
        Me.Controls.Add(Me.codigodoPVP)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.labeldifPR)
        Me.Controls.Add(Me.labelPVPdoPR)
        Me.Controls.Add(Me.PVPdoPR)
        Me.Controls.Add(Me.difPR)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.labelPR2)
        Me.Controls.Add(Me.labelPR1)
        Me.Controls.Add(Me.PR2)
        Me.Controls.Add(Me.PR1)
        Me.Controls.Add(Me.labeldifPVP)
        Me.Controls.Add(Me.labelPR)
        Me.Controls.Add(Me.PRdoPVP)
        Me.Controls.Add(Me.difPVP)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.labelPort1RE)
        Me.Controls.Add(Me.labelPort1RG)
        Me.Controls.Add(Me.label100R)
        Me.Controls.Add(Me.labelRG)
        Me.Controls.Add(Me.labelRE)
        Me.Controls.Add(Me.labelPVP2)
        Me.Controls.Add(Me.labelPVP1)
        Me.Controls.Add(Me.PVP2)
        Me.Controls.Add(Me.pvp1)
        Me.Name = "EP"
        Me.Text = "PrescAvi 3.0 - E. P. / PR"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents labelPort1RE As System.Windows.Forms.Label
    Friend WithEvents labelPort1RG As System.Windows.Forms.Label
    Friend WithEvents label100R As System.Windows.Forms.Label
    Friend WithEvents labelRG As System.Windows.Forms.Label
    Friend WithEvents labelRE As System.Windows.Forms.Label
    Friend WithEvents labelPVP2 As System.Windows.Forms.Label
    Friend WithEvents labelPVP1 As System.Windows.Forms.Label
    Friend WithEvents PVP2 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents pvp1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents difPVP As System.Windows.Forms.MaskedTextBox
    Friend WithEvents PRdoPVP As System.Windows.Forms.MaskedTextBox
    Friend WithEvents labelPR As System.Windows.Forms.Label
    Friend WithEvents labeldifPVP As System.Windows.Forms.Label
    Friend WithEvents labeldifPR As System.Windows.Forms.Label
    Friend WithEvents labelPVPdoPR As System.Windows.Forms.Label
    Friend WithEvents PVPdoPR As System.Windows.Forms.MaskedTextBox
    Friend WithEvents difPR As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents labelPR2 As System.Windows.Forms.Label
    Friend WithEvents labelPR1 As System.Windows.Forms.Label
    Friend WithEvents PR2 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents PR1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents codigodoPVP As System.Windows.Forms.TextBox
    Friend WithEvents codigodoPR As System.Windows.Forms.TextBox
End Class
