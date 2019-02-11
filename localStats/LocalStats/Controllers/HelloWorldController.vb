Imports System.Web.Mvc

Namespace Controllers
    Public Class HelloWorldController
        Inherits Controller

        ' GET: HelloWorld
        Function Index() As ActionResult
            Return View()
        End Function

        Function Welcome() As ActionResult
            Return View()
        End Function

    End Class
End Namespace