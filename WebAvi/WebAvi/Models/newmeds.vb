Public Class newmeds
    Public Property code As Integer
    Public Property nome As String
    Public Property dci As String
    Public Property forma As String
    Public Property dose As String
    Public Property qty As String
    Public Property comp As Integer
    Public Property gh As Integer
    Public Property pvp As Double
    Public Property pr As Double
    Public Property gen As Boolean
    Public Property lab As String
    Public Property alzheimer As Boolean
    Public Property chron As Boolean
    Public Property psicose As Boolean
    Public Property doroncol As Boolean
    Public Property dorcron As Boolean
    Public Property procriacao As Boolean
    Public Property artrite As Boolean
    Public Property top5 As Double
    Public Property dci_obr As Boolean
    Public Property psoriase As Boolean
    Public Property pvpold As Double
    Public Property trocamarca As Boolean
    Public Property pvpmenos3 As Double
    Public Property cnpem As Integer
    Public Property pvpmenos4 As Double
    Public Property pma As Double
    Public Property pvpmenos5 As Double
    Public Property pvpmenos6 As Double


    Public Class prescrito
        Inherits newmeds
    End Class

    Public Class aviado
        Inherits newmeds
    End Class

End Class
