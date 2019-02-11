Namespace Models


    Public Class resultado
        Public Property msg As String
        Public Property alinea As Char
        Public Property excep As Boolean
        Public Property top5 As Boolean
        Public Property cor As String
        Public Property port As Boolean
        Public Property qualport As String
        Public Property pvp As Double
        Public Property sns As Double

        Public Property SelectionStart As String
        Public Property SelectionLength As String
        Public Property SelectionColor As String
        Public Property SelectionbackColor As String

        Public Property ExcepBoxtext As String
        Public Property ExcepBoxtSelectionStart As String
        Public Property ExcepBoxtSelectionLength As String
        Public Property ExcepBoxtSelectionColor As String
        Public Property ExcepBoxtSelectionBackColor As String

        Public Property top5Boxtext As String
        Public Property top5BoxtSelectionStart As String
        Public Property top5BoxtSelectionLength As String
        Public Property top5BoxtSelectionColor As String
        Public Property top5BoxtSelectionBackColor As String

        Public Property verificador As Short

        Public Property cruz1 As Short
        Public Property cruz2 As Short
        Public Property cruz3 As Short
        Public Property cruz4 As Short
    End Class

    Public Class geral
        Inherits resultado
    End Class

    Public Class linha
        Inherits resultado
    End Class



End Namespace

