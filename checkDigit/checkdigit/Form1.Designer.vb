<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.numerotxtbx = New System.Windows.Forms.TextBox()
        Me.aviadotxtbx = New System.Windows.Forms.TextBox()
        Me.verifbtn = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'numerotxtbx
        '
        Me.numerotxtbx.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.numerotxtbx.Location = New System.Drawing.Point(56, 42)
        Me.numerotxtbx.Name = "numerotxtbx"
        Me.numerotxtbx.Size = New System.Drawing.Size(168, 33)
        Me.numerotxtbx.TabIndex = 0
        '
        'aviadotxtbx
        '
        Me.aviadotxtbx.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.aviadotxtbx.Location = New System.Drawing.Point(56, 120)
        Me.aviadotxtbx.MaxLength = 7
        Me.aviadotxtbx.Name = "aviadotxtbx"
        Me.aviadotxtbx.Size = New System.Drawing.Size(168, 33)
        Me.aviadotxtbx.TabIndex = 1
        '
        'verifbtn
        '
        Me.verifbtn.Font = New System.Drawing.Font("Calibri", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.verifbtn.Location = New System.Drawing.Point(100, 191)
        Me.verifbtn.Name = "verifbtn"
        Me.verifbtn.Size = New System.Drawing.Size(86, 32)
        Me.verifbtn.TabIndex = 2
        Me.verifbtn.Text = "verificar"
        Me.verifbtn.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(56, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(141, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "código a verificar (####.#$)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(56, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "nº de aviados"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.verifbtn)
        Me.Controls.Add(Me.aviadotxtbx)
        Me.Controls.Add(Me.numerotxtbx)
        Me.Name = "Form1"
        Me.Text = "check digit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents numerotxtbx As System.Windows.Forms.TextBox
    Friend WithEvents aviadotxtbx As System.Windows.Forms.TextBox
    Friend WithEvents verifbtn As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
