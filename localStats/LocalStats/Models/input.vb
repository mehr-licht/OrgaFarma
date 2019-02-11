Imports System.Data.Entity
Imports System.ComponentModel.DataAnnotations

Public Class input

    Public Property ID As Integer

    <RegularExpression("^\d{9}(\d{2})?$", ErrorMessage:="Código de utente incorrecto")> _
    Public Property utente As Integer

    <StringLength(7, ErrorMessage:="O código do local tem de ter um U seguido de 6 algarismos")> _
    <RegularExpression("^[uU][0-9]{6}$", ErrorMessage:="Código de local incorrecto")> _
    Public Property local As String

    <RegularExpression("^[0-4]$", ErrorMessage:="quantidade de medicamentos incorrecta")>
        Public Property qty As Integer



End Class

Public Class bdContext
    Inherits DbContext
    Public Property Movies As DbSet(Of input)
End Class
