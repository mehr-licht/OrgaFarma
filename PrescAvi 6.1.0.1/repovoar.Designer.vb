<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class repovoar
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.AcederBdToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InserirNaBdToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportarPvpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportarPrToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportarCompToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImportarNovosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RepovoarnovosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(284, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AcederBdToolStripMenuItem, Me.InserirNaBdToolStripMenuItem, Me.ImportarPvpToolStripMenuItem, Me.ImportarPrToolStripMenuItem, Me.ImportarCompToolStripMenuItem, Me.ImportarNovosToolStripMenuItem, Me.RepovoarnovosToolStripMenuItem})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(33, 20)
        Me.ToolStripMenuItem1.Text = "bd"
        '
        'AcederBdToolStripMenuItem
        '
        Me.AcederBdToolStripMenuItem.Name = "AcederBdToolStripMenuItem"
        Me.AcederBdToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.AcederBdToolStripMenuItem.Text = "aceder bd"
        '
        'InserirNaBdToolStripMenuItem
        '
        Me.InserirNaBdToolStripMenuItem.Name = "InserirNaBdToolStripMenuItem"
        Me.InserirNaBdToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.InserirNaBdToolStripMenuItem.Text = "inserir na bd"
        '
        'ImportarPvpToolStripMenuItem
        '
        Me.ImportarPvpToolStripMenuItem.Name = "ImportarPvpToolStripMenuItem"
        Me.ImportarPvpToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.ImportarPvpToolStripMenuItem.Text = "importar pvp"
        '
        'ImportarPrToolStripMenuItem
        '
        Me.ImportarPrToolStripMenuItem.Name = "ImportarPrToolStripMenuItem"
        Me.ImportarPrToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.ImportarPrToolStripMenuItem.Text = "importar pr"
        '
        'ImportarCompToolStripMenuItem
        '
        Me.ImportarCompToolStripMenuItem.Name = "ImportarCompToolStripMenuItem"
        Me.ImportarCompToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.ImportarCompToolStripMenuItem.Text = "importar comp"
        '
        'ImportarNovosToolStripMenuItem
        '
        Me.ImportarNovosToolStripMenuItem.Name = "ImportarNovosToolStripMenuItem"
        Me.ImportarNovosToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.ImportarNovosToolStripMenuItem.Text = "importar novos"
        '
        'RepovoarnovosToolStripMenuItem
        '
        Me.RepovoarnovosToolStripMenuItem.Name = "RepovoarnovosToolStripMenuItem"
        Me.RepovoarnovosToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.RepovoarnovosToolStripMenuItem.Text = "repovoar (novos)"
        '
        'repovoar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "repovoar"
        Me.Text = "repovoar"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AcederBdToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InserirNaBdToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportarPvpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportarPrToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportarCompToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportarNovosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RepovoarnovosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
