<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class cnpem
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
        Me.cnpem_code = New System.Windows.Forms.Label()
        Me.cnpem_dci = New System.Windows.Forms.TextBox()
        Me.cnpem_dose = New System.Windows.Forms.TextBox()
        Me.cnpem_forma = New System.Windows.Forms.TextBox()
        Me.cnpem_qty = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cnpem_but = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cnpem_code
        '
        Me.cnpem_code.AccessibleDescription = "cnpem_code"
        Me.cnpem_code.AccessibleName = "cnpem_code"
        Me.cnpem_code.AutoSize = True
        Me.cnpem_code.Location = New System.Drawing.Point(12, 9)
        Me.cnpem_code.Name = "cnpem_code"
        Me.cnpem_code.Size = New System.Drawing.Size(0, 13)
        Me.cnpem_code.TabIndex = 0
        '
        'cnpem_dci
        '
        Me.cnpem_dci.Location = New System.Drawing.Point(37, 41)
        Me.cnpem_dci.Name = "cnpem_dci"
        Me.cnpem_dci.Size = New System.Drawing.Size(235, 20)
        Me.cnpem_dci.TabIndex = 1
        '
        'cnpem_dose
        '
        Me.cnpem_dose.Location = New System.Drawing.Point(51, 82)
        Me.cnpem_dose.Name = "cnpem_dose"
        Me.cnpem_dose.Size = New System.Drawing.Size(100, 20)
        Me.cnpem_dose.TabIndex = 2
        '
        'cnpem_forma
        '
        Me.cnpem_forma.Location = New System.Drawing.Point(51, 119)
        Me.cnpem_forma.Name = "cnpem_forma"
        Me.cnpem_forma.Size = New System.Drawing.Size(221, 20)
        Me.cnpem_forma.TabIndex = 3
        '
        'cnpem_qty
        '
        Me.cnpem_qty.Location = New System.Drawing.Point(81, 161)
        Me.cnpem_qty.Name = "cnpem_qty"
        Me.cnpem_qty.Size = New System.Drawing.Size(100, 20)
        Me.cnpem_qty.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AccessibleDescription = ""
        Me.Label1.AccessibleName = ""
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 89)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(33, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "dose:"
        '
        'Label2
        '
        Me.Label2.AccessibleDescription = ""
        Me.Label2.AccessibleName = ""
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "quantidade:"
        '
        'Label3
        '
        Me.Label3.AccessibleDescription = ""
        Me.Label3.AccessibleName = ""
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "dci:"
        '
        'Label4
        '
        Me.Label4.AccessibleDescription = ""
        Me.Label4.AccessibleName = ""
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 126)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "forma:"
        '
        'cnpem_but
        '
        Me.cnpem_but.Location = New System.Drawing.Point(104, 208)
        Me.cnpem_but.Name = "cnpem_but"
        Me.cnpem_but.Size = New System.Drawing.Size(75, 23)
        Me.cnpem_but.TabIndex = 5
        Me.cnpem_but.Text = "aceitar"
        Me.cnpem_but.UseVisualStyleBackColor = True
        '
        'cnpem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 251)
        Me.Controls.Add(Me.cnpem_but)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cnpem_qty)
        Me.Controls.Add(Me.cnpem_forma)
        Me.Controls.Add(Me.cnpem_dose)
        Me.Controls.Add(Me.cnpem_dci)
        Me.Controls.Add(Me.cnpem_code)
        Me.Name = "cnpem"
        Me.Text = "cnpem"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cnpem_code As System.Windows.Forms.Label
    Friend WithEvents cnpem_dci As System.Windows.Forms.TextBox
    Friend WithEvents cnpem_dose As System.Windows.Forms.TextBox
    Friend WithEvents cnpem_forma As System.Windows.Forms.TextBox
    Friend WithEvents cnpem_qty As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cnpem_but As System.Windows.Forms.Button
End Class
