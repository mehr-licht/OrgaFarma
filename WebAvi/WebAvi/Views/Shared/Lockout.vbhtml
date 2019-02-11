@Imports System.Web.Mvc
@ModelType HandleErrorInfo
@Code
    ViewBag.Title = "Suspensão"
End Code

<hgroup>
    <h1 class="text-danger">Suspesão.</h1>
    <h2 class="text-danger">Esta conta foi suspensa, por favor tente mais tarde.</h2>
</hgroup>
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