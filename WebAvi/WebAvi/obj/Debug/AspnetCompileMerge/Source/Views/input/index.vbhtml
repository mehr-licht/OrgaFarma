


<script type="text/javascript" src="http://code.jquery.com/jquery-1.7.0.min.js"></script>

<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js">
</script>
@Scripts.Render("~/Scripts/JScript.js")
@Scripts.Render("~/Scripts/jquery.tooltipster.min.js")
@Scripts.Render("~/Scripts/slide.js")
@Scripts.Render("~/Scripts/modernizr.custom.js")
@Scripts.Render("~/Scripts/classie.js")
@Styles.Render("~/Content/bootstrap.js")
@Styles.Render("~/Content/styles.css")
@Styles.Render("~/Content/Default.css")
@Styles.Render("~/Content/component.css")
<style>
#radiado, #pvp1, #pvp2, #pvp3, #pvp4, #sns1, #sns2, #sns3, #sns4, #port1, #port2, #port3, #port4{
    visibility:hidden !important;
}
</style>

<body class="cbp-spmenu-push">
    <nav class="cbp-spmenu cbp-spmenu-vertical cbp-spmenu-left" id="cbp-spmenu-s1">
        <a href="#" id="dia0" class="dias">0</a>
        <a href="#" id="dia00" class="dias">0</a>

        <a href="#" id="dia01" class="dias">01</a>
        <a href="#" id="dia02" class="dias">02</a>
        <a href="#" id="dia03" class="dias">03</a>
        <a href="#" id="dia04" class="dias">04</a>
        <a href="#" id="dia05" class="dias">05</a>
        <a href="#" id="dia06" class="dias">06</a>
        <a href="#" id="dia07" class="dias">07</a>
        <a href="#" id="dia08" class="dias">08</a>
        <a href="#" id="dia09" class="dias">09</a>
        <a href="#" id="dia10" class="dias">10</a>
        <a href="#" id="dia11" class="dias">11</a>
        <a href="#" id="dia12" class="dias">12</a>
        <a href="#" id="dia13" class="dias">13</a>
        <a href="#" id="dia14" class="dias">14</a>
        <a href="#" id="dia15" class="dias">15</a>
        <a href="#" id="dia16" class="dias">16</a>
        <a href="#" id="dia17" class="dias">17</a>
        <a href="#" id="dia18" class="dias">18</a>
        <a href="#" id="dia19" class="dias">19</a>
        <a href="#" id="dia20" class="dias">20</a>
        <a href="#" id="dia21" class="dias">21</a>
        <a href="#" id="dia22" class="dias">22</a>
        <a href="#" id="dia23" class="dias">23</a>
        <a href="#" id="dia24" class="dias">24</a>
        <a href="#" id="dia25" class="dias">25</a>
        <a href="#" id="dia26" class="dias">26</a>
        <a href="#" id="dia27" class="dias">27</a>
        <a href="#" id="dia28" class="dias">28</a>
        <a href="#" id="dia29" class="dias">29</a>
        <a href="#" id="dia30" class="dias">30</a>
        <a href="#" id="dia31" class="dias">31</a>

    </nav>

    <nav class="cbp-spmenu cbp-spmenu-vertical cbp-spmenu-right" id="cbp-spmenu-s2">
        <a href="#" id="labelcnpem0" >0</a>
        <a href="#" id="labelcnpem0">0</a>
        
        @Html.TextBox("cnpemquery", "", New With {.id = "cnpemquery", .class = "textbox", .autocomplete = "off", .minlength = "7", .maxlength = "8", .size = "8", .onkeyup = "querycnpem()"})
        <a href="#" id="labelcnpem1" class="cnpem"></a>
        <a href="#" id="labelcnpem2" class="cnpem"></a>
        <a href="#" id="labelcnpem3" class="cnpem"></a>
        <a href="#" id="labelcnpem4" class="cnpem"></a>
        <a href="#" id="labelcnpem5" class="cnpem"></a>
        <a href="#" id="labelcnpem6" class="cnpem"></a>
        <a href="#" id="labelcnpem7" class="cnpem"></a>
        <a href="#" id="labelcnpem8" class="cnpem"></a>
        <a href="#" id="labelcnpem9" class="cnpem"></a>
                <a href="#" id="labelcnpem10" class="cnpem"></a>
        <a href="#" id="labelcnpem11" class="cnpem"></a>
        <a href="#" id="labelcnpem12" class="cnpem"></a>
        <a href="#" id="labelcnpem13" class="cnpem"></a>
        <a href="#" id="labelcnpem14" class="cnpem"></a>
        <a href="#" id="labelcnpem15" class="cnpem"></a>
        <a href="#" id="labelcnpem16" class="cnpem"></a>
        <a href="#" id="labelcnpem17" class="cnpem"></a>
        <a href="#" id="labelcnpem18" class="cnpem"></a>
        <a href="#" id="labelcnpem19" class="cnpem"></a>
        <a href="#" id="labelcnpem10" class="cnpem"></a>
        <a href="#" id="labelcnpem21" class="cnpem"></a>
        <a href="#" id="labelcnpem22" class="cnpem"></a>
        <a href="#" id="labelcnpem23" class="cnpem"></a>
        <a href="#" id="labelcnpem24" class="cnpem"></a>
        <a href="#" id="labelcnpem25" class="cnpem"></a>
        <a href="#" id="labelcnpem26" class="cnpem"></a>
        <a href="#" id="labelcnpem27" class="cnpem"></a>
        <a href="#" id="labelcnpem28" class="cnpem"></a>
        <a href="#" id="labelcnpem29" class="cnpem"></a>
        <a href="#" id="labelcnpem30" class="cnpem"></a>
    </nav>


    <nav class="cbp-spmenu cbp-spmenu-horizontal cbp-spmenu-bottom" id="cbp-spmenu-s4">
        <div class="container">
            <div class="wrapper">
                <div id="colr">
                    R
                </div>
                <div id="cold">
                    desp
                </div>
                <div id="coluna">

                </div>
                <div class="cabecalho">
                    sávida
                </div>
                <div class="cabecalho">
                    ctt
                </div>
                <div class="cabecalho">
                    samsQ (25)
                </div>
                <div class="cabecalho">
                    sams s.i.
                </div>
                <div class="cabecalho">
                    sams ct
                </div>
                <div class="cabecalho">
                    sams nt
                </div>
                <div class="cabecalho">
                    Médis
                </div>
                <div class="cabecalho">
                    CGD (13)
                </div>
                <div class="cabecalho">
                    1034/09
                </div>
                <div class="cabecalho">
                    SIB (w0)
                </div>
                <div class="cabecalho">
                    mc1
                </div>
                <div class="cabecalho">
                    mc2
                </div>
                <div class="cabecalho">
                    mc3
                </div>
                <div class="cabecalho">
                    mc4
                </div>
                <div class="cabecalho">
                    mc5
                </div>
                <div class="cabecalho">
                    mc6
                </div>
                <div class="cabecalho">
                    mc7
                </div>
            </div>
            <div class="wrapper line1">
                <div class="colr">

                </div>
                <div class="cold">

                </div>
                <div class="coluna">
                    01
                </div>
                <div class="linha1">
                    aa
                </div>
                <div class="linha1">
                    jc
                </div>
                <div class="linha1">
                    o1
                </div>
                <div class="linha1">
                    bv
                </div>
                <div class="linha1">
                    m1
                </div>
                <div class="linha1">
                    j1
                </div>
                <div class="linha1">
                    f1
                </div>
                <div class="linha1">
                    r1
                </div>
                <div class="linha1">
                    sf
                </div>
                <div class="linha1">
                    w1
                </div>
                <div class="linha1">
                    x1
                </div>
                <div class="linha1">
                    x5
                </div>
                <div class="linha1">
                    x9
                </div>
                <div class="linha1">
                    xd
                </div>
                <div class="linha1">xo</div>
                <div class="linha1">xv</div>

                <div class="linha1">xx</div>

            </div>
            <div class="wrapper line2">
                <div class="colr">
                    *
                </div>
                <div class="cold">

                </div>
                <div class="coluna" >
                    48
                </div>
                <div class="linha2">
                    ac
                </div>
                <div class="linha2">
                    je
                </div>
                <div class="linha2">
                    o3
                </div>
                <div class="linha2">
                    by
                </div>
                <div class="linha2">
                    m7
                </div>
                <div class="linha2">
                    j7
                </div>
                <div class="linha2">
                    f7
                </div>
                <div class="linha2">
                    r3
                </div>
                <div class="linha2">
                    sh
                </div>
                <div class="linha2">
                    w3
                </div>
                <div class="linha2">
                    x3
                </div>
                <div class="linha2">
                    x7
                </div>
                <div class="linha2">
                    xb
                </div>
                <div class="linha2">
                    xf
                </div>
                <div class="linha2">
                    xq
                </div>
                <div class="linha2">

                </div>
                <div class="linha2">

                </div>
            </div>
            <div class="wrapper line3">
                <div class="colr">

                </div>
                <div class="cold">
                    *
                </div>
                <div class="coluna" >
                    45
                </div>
                <div class="linha3">
                    ab
                </div>
                <div class="linha3">
                    jd
                </div>
                <div class="linha3">
                    o2
                </div>
                <div class="linha3">
                    bx
                </div>
                <div class="linha3">
                </div>
                <div class="linha3">
                    j4
                </div>
                <div class="linha3">
                    f4
                </div>
                <div class="linha3">
                    r2
                </div>
                <div class="linha3">
                    sg
                </div>
                <div class="linha3">
                    w2
                </div>
                <div class="linha3">
                    x2
                </div>
                <div class="linha3">
                    x6
                </div>
                <div class="linha3">
                    xa
                </div>
                <div class="linha3">
                    xe
                </div>
                <div class="linha3">
                    xp
                </div>
                <div class="linha3">
                    xy
                </div>
                <div class="linha3">
                    xz
                </div>
            </div>
            <div class="wrapper line4">
                <div class="colr">
                    *
                </div>
                <div class="cold">
                    *
                </div>
                <div class="coluna" >
                    49
                </div>
                <div class="linha4">
                    ad
                </div>
                <div class="linha4">
                    jf
                </div>
                <div class="linha4">
                    o4
                </div>
                <div class="linha4">
                    bw
                </div>
                <div class="linha4">
                </div>
                <div class="linha4">
                    j8
                </div>
                <div class="linha4">
                    f8
                </div>
                <div class="linha4">
                    r4
                </div>
                <div class="linha4">
                    si
                </div>
                <div class="linha4">
                    w4
                </div>
                <div class="linha4">
                    x4
                </div>
                <div class="linha4">
                    x8
                </div>
                <div class="linha4">
                    xc
                </div>
                <div class="linha4">
                    xg
                </div>
                <div class="linha4">
                    xr
                </div>
                <div class="linha4">

                </div>
                <div class="linha4">

                </div>
            </div>
            <div class="wrapper line5">
                <div class="colr">
                    ­­/
                </div>
                <div class="cold">
                    prof

                </div>
                <div class="coluna" >
                    41
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                    f2
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">
                </div>
                <div class="linha5">

                </div>
                <div class="linha5">

                </div>
                <div class="linha5">

                </div>
            </div>
            <div class="wrapper line6">
                <div class="colr">
                    /
                </div>
                <div class="cold">
                    4521

                </div>
                <div class="coluna" >
                    42
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                    m3
                </div>
                <div class="linha6">
                    j3
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">
                </div>
                <div class="linha6">

                </div>
                <div class="linha6">

                </div>
                <div class="linha6">

                </div>
            </div>

            <div class="wrapper line7">
                <div class="colr">
                    /
                </div>
                <div class="cold">
                    11387

                </div>
                <div class="coluna" >
                    67
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                    m2
                </div>
                <div class="linha7">
                    j2
                </div>
                <div class="linha7">
                    f9
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">
                </div>
                <div class="linha7">

                </div>
                <div class="linha7">

                </div>
                <div class="linha7">

                </div>
            </div>

            <div class="wrapper line8">
                <div class="colr">
                    /

                </div>
                <div class="cold">
                    manip
                </div>
                <div class="coluna" >
                    47
                </div>
                <div class="linha8">
                    aj
                </div>
                <div class="linha8">
                    cd
                </div>
                <div class="linha8">
                    o9
                </div>
                <div class="linha8">
                    bz
                </div>
                <div class="linha8">
                </div>
                <div class="linha8">
                    j9
                </div>
                <div class="linha8">
                    fc
                </div>
                <div class="linha8">
                    r9
                </div>
                <div class="linha8">
                    sj
                </div>
                <div class="linha8">
                    w9
                </div>
                <div class="linha8">
                    xh
                </div>
                <div class="linha8">
                </div>
                <div class="linha8">
                </div>
                <div class="linha8">
                </div>
                <div class="linha8">
                    xs

                </div>
                <div class="linha8">
                    xk

                </div>
                <div class="linha8">

                </div>
            </div>


        </div>
    </nav>
    
    


        <br /><br />
    <div id="radiado">
        <input type="radio" id="radio01" name="radios" value="01" checked>
        <label for="radio01">01</label>
        <input type="radio" id="radio48" name="radios" value="48">
        <label for="radio48">48</label>
        <input type="radio" id="radio45" name="radios" value="45">
        <label for="radio45">45</label>
        <input type="radio" id="radio49" name="radios" value="49">
        <label for="radio49">49</label>
        <input type="radio" id="radio67" name="radios" value="67">
        <label for="radio67">67</label>
        <input type="radio" id="radio42" name="radios" value="42">
        <label for="radio42">42</label>
        <input type="radio" id="radioDS" name="radios" value="DS">
        <label for="radioDS">DS</label>
    </div>
        <div Class="container">

           <br/>


            <form id="inputform" Class="form-large" onreset="limpar()">

                <div Class="col-md-12 col-xs-12">
                    <div Class="col-md-6 col-xs-6 esquerda">
                        <div Class="col-md-4 col-xs-4 espaco">

                            <Label for="prescrito1" id="label1"></Label>

                        </div>
                        <div Class="col-md-4 col-xs-4 prescritos">

                            @Html.TextBox("p1", "", New With {.id = "p1", .class = "textbox", .tabindex = "1", .autocomplete = "off", .minlength = "7", .maxlength = "8", .size = "8", .required = "required", .autofocus = "autofocus", .onfocus = "focusFunction()"})<!--.onkeyup = "movetoNext(this, 'p2')", -->



                        </div>
                    </div>

                    <div Class="col-md-6 col-xs-6 direita">
                        <div Class="col-md-4 col-xs-4 espaco">
                            <Label for="aviado1" id="label2"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 aviados">
                            @Html.TextBox("a1", "", New With {.id = "a1", .class = "textbox", .tabindex = "5", .autocomplete = "off", .minlength = "7", .maxlength = "7", .onkeyup = "movetoNext(this, 'a2')", .size = "7", .required = "required", .onfocus = "focusFunction()"})<!-- -->
                        </div>

                        <div Class="col-md-4 col-xs-4 labellinha">
                            <Label for="aviado1" id="labellinha1" class="labellinhas"></Label>
                        </div>
                    </div>
                </div>



                <div Class="row">
                    <div class="col-md-12 col-md-12">
                        <div Class="col-md-6 col-xs-6 esquerda">
                            <div Class="col-md-4 col-xs-4 espaco"></div>
                            <div Class="col-md-4 col-xs-4 prescritos">
                                <div Class="col-md-4 col-xs-4 dadospr">
                                    <select id="pr1" name="pr1" class="prlinhas" style="visibility: hidden"> </select></div>
                               
                            </div>
                        </div>

                        <div Class="col-md-6 col-xs-6 direita">
                            <div class="col-md-4 col-xs-4 dadosport">
                                <select id="port1" name="port1" class="portlinhas" style="visibility: hidden" onchange="showchange(1)">  </select>
                            </div>

                            <div Class="col-md-4 col-xs-4 aviados">
                                <div Class="col-md-2 col-xs-2 selectpvp">
                                    <select id="pvp1" name="pvp1" class="pvplinhas" style="visibility: hidden"  onchange="showchange(1)">
                                        @*<option value="option11"></option>*@

                                    </select>
                                </div>
                                <div Class="col-md-2 col-xs-2">
                                    <Label for="aviado1sns" id="sns1" class="snslinhas" style="visibility: hidden"></Label>
                                </div>
                            </div>
                            <div Class="col-md-4 col-xs-4 labellinha"></div>
                        </div>
                    </div>
                </div>




                <div Class="col-md-12 col-xs-12">
                    <div Class="col-md-6 col-xs-6 esquerda">
                        <div Class="col-md-4 col-xs-4 espaco">

                            <Label for="prescrito2" id="label3"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 prescritos">
                            @Html.TextBox("p2", "", New With {.id = "p2", .class = "textbox", .tabindex = "2", .autocomplete = "off", .minlength = "7", .maxlength = "8", .size = "8", .onfocus = "focusFunction()"})<!--.onkeyup = "movetoNext(this, 'p3')",-->

                        </div>
                    </div>

                    <div Class="col-md-6 col-xs-6 direita">
                        <div Class="col-md-4 col-xs-4 espaco">
                            <Label for="aviado2" id="label4"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 aviados">
                            @Html.TextBox("a2", "", New With {.id = "a2", .class = "textbox", .tabindex = "6", .autocomplete = "off", .minlength = "7", .maxlength = "7", .onkeyup = "movetoNext(this, 'a3')", .size = "7", .onfocus = "focusFunction()"})<!---->
                        </div>
                        <div Class="col-md-4 col-xs-4 labellinha">
                            <Label for="aviado2" id="labellinha2" class="labellinhas"></Label>
                        </div>
                    </div>
                </div>


                <div Class="row">
                    <div class="col-md-12 col-md-12">
                        <div Class="col-md-6 col-xs-6 esquerda">
                            <div Class="col-md-4 col-xs-4 espaco"></div>
                            <div Class="col-md-4 col-xs-4 prescritos">
                                <div Class="col-md-2 col-xs-2 dadospr">
                                    <select id="pr2" name="pr2" class="prlinhas" style="visibility: hidden"> </select>
                                </div>

                            </div>
                        </div>

                        <div Class="col-md-6 col-xs-6 direita">
                            <div class="col-md-4 col-xs-4 dadosport">
                                <select id="port2" name="port2" class="portlinhas" style="visibility: hidden" onchange="showchange(2)">  </select>
                            </div>

                            <div Class="col-md-4 col-xs-4 aviados">
                                <div Class="col-md-2 col-xs-2 selectpvp">
                                    <select id="pvp2" name="pvp2" class="pvplinhas" style="visibility: hidden" onchange="showchange(2)">
                                        @*<option value="option21"></option>*@

                                    </select>
                                </div>
                                <div Class="col-md-2 col-xs-2">
                                    <Label for="aviado2sns" id="sns2" class="snslinhas" style="visibility: hidden"></Label>
                                </div>
                            </div>
                            <div Class="col-md-4 col-xs-4 labellinha"></div>
                        </div>
                    </div>
                </div>

                <div Class="col-md-12 col-xs-12">
                    <div Class="col-md-6 col-xs-6 esquerda">
                        <div Class="col-md-4 col-xs-4 espaco">

                            <Label for="prescrito3" id="label5"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 prescritos">
                            @Html.TextBox("p3", "", New With {.id = "p3", .class = "textbox", .tabindex = "3", .autocomplete = "off", .minlength = "7", .maxlength = "8", .size = "8", .onfocus = "focusFunction()"})<!-- .onkeyup = "movetoNext(this, 'p4')",-->
                        </div>
                    </div>

                    <div Class="col-md-6 col-xs-6 direita">
                        <div Class="col-md-4 col-xs-4 espaco">
                            <Label for="aviado3" id="label6"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 aviados">
                            @Html.TextBox("a3", "", New With {.id = "a3", .class = "textbox", .tabindex = "7", .autocomplete = "off", .minlength = "7", .maxlength = "7", .onkeyup = "movetoNext(this, 'a4')", .size = "7", .onfocus = "focusFunction()"})<!-- -->
                        </div>
                        <div Class="col-md-4 col-xs-4 labellinha">
                            <Label for="aviado3" id="labellinha3" class="labellinhas"></Label>
                        </div>
                    </div>
                </div>


                <div Class="row">
                    <div class="col-md-12 col-md-12">
                        <div Class="col-md-6 col-xs-6 esquerda">
                            <div Class="col-md-4 col-xs-4 espaco"></div>
                            <div Class="col-md-4 col-xs-4 prescritos">
                                <div Class="col-md-2 col-xs-2 dadospr">
                                    <select id="pr3" name="pr3" class="prlinhas" style="visibility: hidden"></select>
                                </div>

                            </div>
                        </div>

                        <div Class="col-md-6 col-xs-6 direita">
                            <div class="col-md-4 col-xs-4 dadosport">
                                <select id="port3" name="port3" class="portlinhas" style="visibility: hidden" onchange="showchange(3)">  </select>
                            </div>

                            <div Class="col-md-4 col-xs-4 aviados">
                                <div Class="col-md-2 col-xs-2 selectpvp">
                                    <select id="pvp3" name="pvp3" class="pvplinhas" style="visibility: hidden" onchange="showchange(3)">
                                        @*<option value="option31"></option>*@

                                    </select>
                                </div>
                                <div Class="col-md-2 col-xs-2">
                                    <Label for="aviado3sns" id="sns3" class="snslinhas" style="visibility: hidden"></Label>
                                </div>
                            </div>
                            <div Class="col-md-4 col-xs-4 labellinha"></div>
                        </div>
                    </div>
                </div>

                <div Class="col-md-12 col-xs-12">
                    <div Class="col-md-6 col-xs-6 esquerda">
                        <div Class="col-md-4 col-xs-4 espaco">

                            <Label for="prescrito4" id="label7"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 prescritos">
                            @Html.TextBox("p4", "", New With {.id = "p4", .class = "textbox", .tabindex = "4", .autocomplete = "off", .minlength = "7", .maxlength = "8", .size = "8", .onfocus = "focusFunction()"})<!-- .onkeyup = "movetoNext(this, 'a1')", -->
                        </div>
                    </div>

                    <div Class="col-md-6 col-xs-6 direita">
                        <div Class="col-md-4 col-xs-4 espaco">
                            <Label for="aviado4" id="label8"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 aviados">
                            @Html.TextBox("a4", "", New With {.id = "a4", .class = "textbox", .tabindex = "8", .autocomplete = "off", .minlength = "7", .maxlength = "7", .onkeyup = "movetoNext(this, 'btnSubmit')", .size = "7", .onfocus = "focusFunction()"})<!---->
                        </div>
                        <div Class="col-md-4 col-xs-4 labellinha">
                            <Label for="aviado4" id="labellinha4" class="labellinhas"></Label>
                        </div>
                    </div>
                </div>


                <br>
                <div Class="row">
                    <div class="col-md-12 col-md-12">
                        <div Class="col-md-6 col-xs-6 esquerda">
                            <div Class="col-md-4 col-xs-4 espaco"></div>
                            <div Class="col-md-4 col-xs-4 prescritos">
                                <div Class="col-md-2 col-xs-2 dadospr">
                                    <select id="pr4" name="pr4" class="prlinhas" style="visibility: hidden"></select>
                                </div>

                            </div>
                        </div>

                        <div Class="col-md-6 col-xs-6 direita">
                            <div class="col-md-4 col-xs-4 dadosport">
                                <select id="port4" name="port4" class="portlinhas" style="visibility: hidden" onchange="showchange(4)">  </select>
                            </div>

                            <div Class="col-md-4 col-xs-4 aviados">
                                <div Class="col-md-2 col-xs-2 selectpvp">
                                    <select id="pvp4" name="pvp4" class="pvplinhas" style="visibility: hidden" onchange="showchange(4)">
                                        @*<option value="option41"></option>*@

                                    </select>
                                </div>
                                <div Class="col-md-2 col-xs-2">
                                    <Label for="aviado4sns" id="sns4" class="snslinhas" style="visibility: hidden"></Label>
                                </div>
                            </div>
                            <div Class="col-md-4 col-xs-4 labellinha"></div>
                        </div>
                    </div>
                </div>
                <br>

                <div Class="col-md-12 col-xs-12">
                    <div Class="col-md-6 col-xs-6 esquerda">
                        <div Class="col-md-4 col-xs-4 espaco">

                            <Label for="nada1a" id="label9"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4 prescritos">
                            <Label for="nada1b" id="label10"></Label>
                        </div>
                    </div>

                    <div Class="col-md-6 col-xs-6 direita">
                        <div Class="col-md-4 col-xs-4 espaco">
                            <Label for="nada2a" id="label11"></Label>
                        </div>

                        <div Class="col-md-4 col-xs-4">
                            <input id="btnSubmit" type="button"  name="btnSubmit" value="comparar" tabindex="9" class="btn" />
                            <input id="resetbtn" type="reset" class="reset" />
                            <!-- <input type="submit" name="btnSubmit" value="Send" /-->
                            <!-- <asp:Button ID="btnSubmit" Text="Submit" runat="server" ClientIDMode="Static" OnClientClick="return false;" />-->
                        </div>
                    </div>
                </div>







            </form>

        </div>


</body>

@Html.Partial("~/views/prescavi/_mostrar.vbhtml")
