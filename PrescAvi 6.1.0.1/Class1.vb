Class TransparentRichTextBox
    Inherits RichTextBox
    Public Sub New()
        MyBase.ScrollBars = RichTextBoxScrollBars.None
        MyBase.BorderStyle = Windows.Forms.BorderStyle.None
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H20
            Return cp
        End Get
    End Property
    Protected Overloads Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
    End Sub

End Class