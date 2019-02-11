

<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>


@Scripts.Render("~/Scripts/JScript.js")
@Styles.Render("~/Content/styles.css")

<BODY id="cortotal">

    <h2 id="titulomostrar" class="mostrado"></h2>
    <div class="container">
        <div class="row">
            <div class="col-md-offset-1 col-md-2 semaforo">
                @*<div id="semaforo" style="background-image: url(~/img/empty.gif); height: 200px; width: 200px; border: 0px transparent;">*@
                <img src="~/img/empty.gif" id="semaforo" class="mostrado">
                    @*<h4 id="labeltotalports" ></h4>
                    <h4 id="labeltotalexcep" ></h4>
                    <h4 id="labeltotaltop5" ></h4>*@
                @*</div>*@
                
                @**@
            </div>
            <div class="col-md-offset-1 col-md-1 espaco">
            </div>
            <div class="col-md-offset-1 col-md-3 total">
                <h1 id="labeltotal0" class="mostrado"></h1>
                <button class="btn" onclick="resetbtn.click()" id="btnlimpar">limpar</button>
            </div>
            @*<div class="col-md-offset-1 col-md-1 espaco">
                </div>*@
            <div class="col-md-offset-1 col-md-5 linhas">
                <h4 id="labeltotal1" class="totallinhas"></h4>
                <h4 id="labeltotal2" class="totallinhas"></h4>
                <h4 id="labeltotal3" class="totallinhas"></h4>
                <h4 id="labeltotal4" class="totallinhas"></h4>
                @*<h4 id="labeltotal5" class="mostrado"></h4>
                    <h4 id="labeltotal6" class="mostrado"></h4>*@
            </div>
        </div>
    </div>


</BODY>