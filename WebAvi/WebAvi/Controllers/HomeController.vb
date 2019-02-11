Public Class HomeController
    Inherits System.Web.Mvc.Controller

    Function Index() As ActionResult
        Return View()
    End Function

    Function About() As ActionResult
        ViewData("Message") = "Acerca."

        Return View()
    End Function

    Function Contact() As ActionResult
        ViewData("Message") = "Contactos."

        Return View()
    End Function

    <Authorize>
    Function Downloads() As ActionResult
        ViewData("Message") = "Downloads."

        Return View()
    End Function


    Function Terms() As ActionResult
        ViewData("Message") = "Termos de utilização / política de protecção de dados"

        Return View()
    End Function
End Class
