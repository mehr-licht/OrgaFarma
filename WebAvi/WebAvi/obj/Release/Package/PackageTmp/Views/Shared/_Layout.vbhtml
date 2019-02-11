<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>@ViewBag.Title - WebAvi</title>
    @Styles.Render("~/Content/css")
    @Scripts.Render("~/bundles/modernizr")
   @Scripts.Render("~/Scripts/JScript.js")
    @Styles.Render("~/Content/t.css")
</head>
<body>
    <script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>
   
    <div class="navbar navbar-inverse navbar-fixed-top">
        <div class="container">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                @Html.ActionLink("WebAvi", "Index", "Input", New With {.area = ""}, New With {.class = "navbar-brand"})
            </div>
            <div class="navbar-collapse collapse">
                <ul class="nav navbar-nav">
            
                    <li><a href="#" id="showLeft" class="botaotexto">Datas</a></li>
                    <li><a href="#" id="showBottom" class="botaotexto">Organismos</a></li>
                    <li><a href="#" id="showRight" class="botaotexto">Cnpem</a></li>
                    <li><a href="#" id="showpvp" class="botaotexto" onclick="togglepvp()">pvp</a></li>
<li>@Html.ActionLink("Acerca", "Index", "Home")</li>
                   
                </ul>
                @Html.Partial("_LoginPartial")
            </div>
        </div>
    </div>
    <div class="container body-content">
        @RenderBody()
        <hr />
        <footer>
            <p>&copy; @DateTime.Now.Year - OrgaFarma</p>
        </footer>
    </div>

    @Scripts.Render("~/bundles/jquery")
    @Scripts.Render("~/bundles/bootstrap")
    @RenderSection("scripts", required:=False)
</body>
</html>
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