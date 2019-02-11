@Code
    ViewBag.Title = "Confirmar email"
End Code

<h2>@ViewBag.Title.</h2>
<div>
    <p>
        Obrigado por confirmar o email. Por favor @Html.ActionLink("Click aqui para entrar", "Login", "Account", routeValues:=Nothing, htmlAttributes:=New With {Key .id = "loginLink"})
    </p>
</div>
<style>
    BODY {
        background-color: cadetblue;
    }


        body input[class='form-control'] {
            background-color: lightblue;
            color: black;
            font-family: Verdana;
            font-language-override: "PT";
            border: 2px solid #456879;
            border-radius: 10px;
            text-align: center;
        }
</style>