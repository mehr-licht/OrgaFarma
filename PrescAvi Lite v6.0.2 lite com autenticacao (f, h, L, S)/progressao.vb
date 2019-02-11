Public Class progressao

    Private Sub PROGRESSAO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        rtb()
    End Sub

    Sub rtb()

        RichTextBox1.SelectAll()
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

End Class