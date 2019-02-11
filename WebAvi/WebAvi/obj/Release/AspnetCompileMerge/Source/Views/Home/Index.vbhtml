@Code
    ViewData("Title") = "início"
End Code
<br />
<br />
@Html.ActionLink("Acerca", "About", "Home").
<br />
<br />
@Html.ActionLink("Contactos", "Contact", "Home")
<br />
<br />
@Html.ActionLink("Downloads", "Downloads", "Home")



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