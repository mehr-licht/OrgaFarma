Imports System.Data.Entity
Imports System.ComponentModel.DataAnnotations

Namespace Models

    Public Class input
        <Required(ErrorMessage:="É necessário um prescrito")>
        <RegularExpression("^\d{7}(\d{1})?$", ErrorMessage:="CNP de 7 algarismos ou CNPEM de 8 algarismos")>
        Public Property input1 As Integer

        <RegularExpression("^\d{7}(\d{1})?$", ErrorMessage:="CNP de 7 algarismos ou CNPEM de 8 algarismos")>
        Public Property input2 As Integer

        <RegularExpression("^\d{7}(\d{1})?$", ErrorMessage:="CNP de 7 algarismos ou CNPEM de 8 algarismos")>
        Public Property input3 As Integer

        <RegularExpression("^\d{7}(\d{1})?$", ErrorMessage:="CNP de 7 algarismos ou CNPEM de 8 algarismos")>
        Public Property input4 As Integer

        <Required(ErrorMessage:="É necessário um aviado")>
        <RegularExpression("^\d{7}?$", ErrorMessage:="CNP de 7 algarismos")>
        Public Property input5 As Integer

        <RegularExpression("^\d{7}?$", ErrorMessage:="CNP de 7 algarismos")>
        Public Property input6 As Integer

        <RegularExpression("^\d{7}?$", ErrorMessage:="CNP de 7 algarismos")>
        Public Property input7 As Integer

        <RegularExpression("^\d{7}?$", ErrorMessage:="CNP de 7 algarismos")>
        Public Property input8 As Integer
    End Class
End Namespace
