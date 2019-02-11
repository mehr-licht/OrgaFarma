Public Class embalagem

    Public Class ordem
        '(farei a1, a2, a3, a4, p1, p2, p3, p4, ad1, ad2, at, aq, pd1, pd2, pt, pq)

        Public code As Integer
        Public dci As String
        Public nome As String
        Public forma As String
        Public dose As String
        Public qty As String
        Public comp As Short
        Public gh As Short
        Public pvp As Double
        Public pr As Double
        Public gen As Boolean
        Public lab As String
        Public d4250 As Boolean
        Public d1234 As Boolean
        Public d21094 As Boolean
        Public d10279 As Boolean
        Public d10280 As Boolean
        Public d10910 As Boolean
        Public d14123 As Boolean
        Public top5 As Double
        Public dci_obr As Boolean
        Public lei6 As Boolean
        Public pvpmenos1 As Double
        Public trocamarca As Boolean
        Public CNPEM As Integer
        Public duplicado As Boolean
        Public nDCI As Boolean
        Public pvpmenos2 As Double
        Public pvptop5 As Boolean
        Public pvpmenos1top5 As Boolean
        Public pvpmenos2top5 As Boolean
        Public pvpun As Double
        Public pvpmenos1un As Double
        Public pvpmenos2un As Double
        Public excepa As Boolean
        Public excepb As Boolean
        Public excepc As Boolean
        Public prescrito As Integer
        Public cruzacom As Single
        Public result As String
        Public porCNPEM As Boolean
        Public pvpmenos3 As Double
        Public pvpmenos4 As Double
        Public pvpmenos5 As Double
        Public pvpexcepC As Double
        Public arraypvp As Array
    End Class





    Public Class cruzamento
        Public code As Boolean
        Public dci As Boolean
        Public nome As Boolean
        Public forma As Boolean
        Public dose As Boolean
        Public qty As Boolean
        Public comp As Boolean
        Public gh As Boolean
        Public pvp As Boolean
        Public pr As Boolean
        Public gen As Boolean
        Public lab As Boolean
        Public d4250 As Boolean
        Public d1234 As Boolean
        Public d21094 As Boolean
        Public d10279 As Boolean
        Public d10280 As Boolean
        Public d10910 As Boolean
        Public d14123 As Boolean
        Public top5 As Double
        Public dci_obr As Boolean
        Public lei6 As Boolean
        Public pvpmenos1 As Boolean
        Public pvpmenos2 As Boolean
        Public trocamarca As Boolean
        Public CNPEM As Boolean
        Public nDCI As Boolean  '?
        Public excepa As Single 'se 
        Public excepb As Single
        Public excepc As Single
        Public xqty As Single 'menor, 1<x<=1,5, x>1,5
        Public xpvp As Single 'se existem ambos e compara maior, menor
        Public xgh As Single 'se existem ambos
        Public xcnpem As Single 'se existem ambos
        Public xnDCI As Boolean 'se existem ambos
        Public porDCImesmoCNPEM As Boolean
        Public porDCIdifCNPEM As Boolean
        Public porMARCAmesmoCNPEM As Boolean
        Public porMARCAdifCNPEM As Boolean
        Public ranking As String
        Public anulado As Boolean
        Public qualP As Single
        Public marcadifsemhavergens As Single
        Public erro As Boolean
    End Class






End Class
