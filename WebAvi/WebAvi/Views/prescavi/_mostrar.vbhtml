

<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>


@Scripts.Render("~/Scripts/JScript.js")
@Styles.Render("~/Content/styles.css")
<style>
    /*#snstotal{
        visibility: hidden !important;
    }*/
    .semaforo { 
   position: relative; 
   /* width: 100%; for IE 6 */
}

h3 { 
   position: absolute; 
   top: 40px; 
   left: 40px; 
   width: 100%; 
   font-weight: bold;
   color: black;
   text-align:center;
}

</style>
<BODY id="cortotal">

    <h2 id="titulomostrar" class="mostrado"></h2>
    <div class="container">
        <div class="row">
            <div class="col-md-offset-1 col-md-2 semaforo">
                
                <img src="~/img/empty.gif" id="semaforo" class="mostrado">
                    <h3 id="snstotal" ></h3>
            
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
                <h4 id="labeltotal5" class="totallinhas"></h4>
                <h4 id="labeltotal6" class="totallinhas"></h4>
                <h4 id="labeltotal7" class="totallinhas"></h4>
     
            </div>
        </div>
    </div>


</BODY>