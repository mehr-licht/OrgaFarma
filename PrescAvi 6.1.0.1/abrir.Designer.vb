<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class abrir
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
        Me.abrirtextbox = New System.Windows.Forms.TextBox
        Me.abrirbutao = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'abrirtextbox
        '
        Me.abrirtextbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.abrirtextbox.Location = New System.Drawing.Point(63, 29)
        Me.abrirtextbox.Name = "abrirtextbox"
        Me.abrirtextbox.Size = New System.Drawing.Size(100, 31)
        Me.abrirtextbox.TabIndex = 0
        '
        'abrirbutao
        '
        Me.abrirbutao.BackColor = System.Drawing.Color.PaleGreen
        Me.abrirbutao.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.abrirbutao.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.abrirbutao.Location = New System.Drawing.Point(163, 29)
        Me.abrirbutao.Name = "abrirbutao"
        Me.abrirbutao.Size = New System.Drawing.Size(24, 30)
        Me.abrirbutao.TabIndex = 1
        Me.abrirbutao.Text = "√"
        Me.abrirbutao.UseVisualStyleBackColor = False
        '
        'abrir
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(230, 85)
        Me.Controls.Add(Me.abrirbutao)
        Me.Controls.Add(Me.abrirtextbox)
        Me.Name = "abrir"
        Me.Text = "abrir"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents abrirtextbox As System.Windows.Forms.TextBox
    Friend WithEvents abrirbutao As System.Windows.Forms.Button
End Class
