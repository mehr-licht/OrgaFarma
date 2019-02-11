Namespace Passdata
    Public Class InputController
        Inherits System.Web.Mvc.Controller

        '
        ' GET: /Input
        <Authorize>
        Function index() As ActionResult
            Return View()
        End Function

    End Class
End Namespace