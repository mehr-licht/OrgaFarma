var organterior = "01";
var mudado = "00000";
var pvpstatus = false;
$('#showpvp').attr('style', 'color: white');


function togglepvp() {
    
    pvpstatus ? pvpstatus = false : pvpstatus = true;
    
    if (pvpstatus == false) {
   
        $('#showpvp').attr('style', 'color: grey');
     
        $('#radiado').attr('style', 'visibility: hidden !important');
  
        for (hid = 1; hid < 5; hid++) {
            $('#pvp' + hid).attr('style', 'visibility: hidden !important');
            $('#sns' + hid).attr('style', 'visibility: hidden !important');
            $('#port' + hid).attr('style', 'visibility: hidden !important');

        }
    } else {
 
        $('#showpvp').attr('style', 'color: white');
   
        $('#radiado').css("visibility", "visible");
    
        for (hid = 1; hid < 5; hid++) {
            $('#pvp' + hid).css("visibility", "hidden");
            $('#sns' + hid).css("visibility", "hidden");
            $('#port' + hid).css("visibility", "hidden");
        }
    };
};



function tooltipalinea(p, a) {
    switch (a) {
        case "f":
            document.getElementById(p).setAttribute('title', "apresentação diferente");
            break;
        case "g":
            document.getElementById(p).setAttribute('title', "dosagem diferente");
            break;
        case "h":
            document.getElementById(p).setAttribute('title', "caixa maior que a prescrita");
            break;
        case "k":
            document.getElementById(p).setAttribute('title', "embalagem repetida");
            break;
        case "l":
            document.getElementById(p).setAttribute('title', "caixa menor que a prescrita");
            break;
        case "s":
            document.getElementById(p).setAttribute('title', "marca diferente sem haver genéricos");
            break;
        case "w":
            document.getElementById(p).setAttribute('title', "CNPEM diferente");
            break;
        case "y":
            document.getElementById(p).setAttribute('title', "dci diferente / medicamento não prescrito");
            break;
        case "z":
            document.getElementById(p).setAttribute('title', "medicamento não prescrito / dci diferente");
            break;
    }
}

$("#btnSubmit").on({
    click: function () {
        document.getElementById("btnSubmit").focus();
    }
});

$("input[type=text]").keydown(function (e) {
    
            if (e.which === 32){
                return false;
        }
            if (e.which === 27) {
                limpar();
            }


    if (e.which == 13 || e.which == 9) {
        var organismo = jQuery('input[name=radios]:checked').val();
        //switch (e) {
        //    case 9:
        //        var focado = parseFloat(document.getElementById(document.activeElement.id))-1;
        //        break;
        //    case 13:
        //        var focado = document.getElementById(document.activeElement.id);
        //        break;
        //}
        var focado = document.getElementById(document.activeElement.id);
        var fim = focado.id.substring(1);
        var inicio = focado.id.substring(0, 1);
        var fimseg = fim;
        var anterior = inicio + (fim - 1);
        var seguinte = inicio + (parseFloat(fim) + 1);
        var primeiro = focado.value.substring(0, 1);
        var primeiros3 = focado.value.substring(0, 3);
   
        if (seguinte == "p5") {
            seguinte = "a1";
        }
        if (focado.value == " ") {//se focado é space
            document.getElementById(focado.id).value = ""; //limpa o focado vazio 
        }

      
        if (!(primeiro == " " || ((focado.value.length == 8) && (primeiro != "5")) || ((organismo == "DS") && (primeiros3 != "619")) || ((organismo != "DS") && (primeiros3 == "619")))) {//negar as novas condições
            
            if (focado.value.length == 0) { //enter no vazio
                if (fim != 1) { //se não é p1 nem a1
                    if (inicio == "a" && fim != "1") {
                        document.getElementById("btnSubmit").style.visibility = "visible"; //se enter no a1 comparar fica visivel
                    }


                    if (document.getElementById(anterior).value.length == 0) { //se anterior vazio
                        document.getElementById(anterior).focus(); //foca no anterior

                    } else { //se anterior não vazio
                        if (seguinte == "a5") { //se é a4

                            document.getElementById("btnSubmit").click(); //foca no seguinte

                        } else { //se nºao é a4
                            switch (inicio) {
                                case "p":
                                    document.getElementById("a1").focus(); //enter no p vazio  <> p1 vai focar no a1
                                    break;
                                case "a":
                                    document.getElementById("btnSubmit").click(); //enter no a vazio  <> a1 vai clickar botão
                            }


                        }
                    }
                } else {

                    document.getElementById(focado.id).focus(); //foca no próprio

                }
            } else { //enter em não vazio
                if (inicio == "a" && fim != "1") {
                    document.getElementById("btnSubmit").style.visibility = "visible"; //se enter no a1 comparar fica visivel
                }
                
                if (focado.value.length < 7 || (isNaN(focado.value))) { //focado incompleto ou com simbolos/letras
                    if (isNaN(focado.value) || primeiro == " ") {
                        document.getElementById(focado.id).value = ""; //limpa o focado vazio
                    }
                    if (anterior.value.length == 0) {
                        document.getElementById(anterior).focus(); //foca no anterior
                        document.getElementById(anterior).trigger(e); //manda enter para voltar a testar anterior
                    } else {
                        document.getElementById(focado.id).focus(); //foca no mesmo

                    }
                } else {
                    if (fim != 1) { //não é p1 nem a1
                        if (document.getElementById(anterior).value.length == 0) { //anterior vazio
                            document.getElementById(anterior).focus(); //foca no anterior

                        } else {
                            if (seguinte == "a5") { //está no a4

                                document.getElementById("btnSubmit").click(); //click o submit

                            } else { //não está no a4
                                document.getElementById(seguinte).focus(); //foca no seguinte

                            }
                        }
                    } else { //é p1 ou a1
                        document.getElementById(seguinte).focus(); //foca no seguinte p2 ou a2
                        if (inicio == "a") {
                            document.getElementById("btnSubmit").style.visibility = "visible"; //se enter no a1 comparar fica visivel
                        }
                    }
                }

            }


        } else {
         
            if ((((organismo == "DS") && (primeiros3 != "619")) || ((organismo != "DS") && (primeiros3 == "619")))) {
                document.getElementById(focado.id).value = "diabetes";
            } else {

                document.getElementById(focado.id).value = ""; //limpa o focado vazio
            }

            if (anterior.value.length == 0) {
                 document.getElementById(anterior).focus(); //foca no anterior
                 document.getElementById(anterior).trigger(e); //manda enter para voltar a testar anterior
            } else {
                 document.getElementById(focado.id).focus(); //foca no mesmo

           
            }

        }




        //if(inicio=="p"){
        //    if (document.getElementById(focado.id).value >= 60000000 || (document.getElementById(focado.id).value > 10000000 && document.getElementById(focado.id).value < 50000000)) {
        //        document.getElementById(focado.id).style.backgroundColor = "red";
        //    }
        //}

    } //não é enter

}); //fim função






$(document).ready(function () {
    var organterior = jQuery('input[name=radios]:checked').val();
    

    $("input[name=radios]:radio").change(function () {

        var organismo = jQuery('input[name=radios]:checked').val();

        if (organismo != organterior) {

            var last9 = ($('#semaforo').attr('src').slice(-9));

            if (last9 != "empty.gif") { //se mostrado 
            
                for (ll = 1; ll < 5; ll++) {
                   
                        if (mudado != organismo + "" + $('#pvp' + ll).children('option:selected').index() + "" + $('#port' + ll).children('option:selected').index() + "" + ll) {
                          
                            if ($('#labellinha' + ll).text() != "") { //se linha com resultados
                          
                                if ((organterior != "45" && organterior != "49") && (organismo == "45" || organismo == "49")) {
                                    for (ss = 10; ss >= 0; ss--) {
                                        for (rr = 1; rr < 5; rr++) {
                                            if (document.getElementById("port" + rr).length >= ss) {
                                                document.getElementById("port" + rr).remove(ss);
                                            }
                                        }
                                    }

                                    if ($('#port' + ll).children('option').length > 1) {
                                        
                                        
                                        var newd = $('#labellinha' + ll).text().search("Despacho");
                                        var newl = $('#labellinha' + ll).text().search("Lei");
                                        newd > newl ? newp = newd : newp = newl;
                                        var newstring = $('#labellinha' + ll).text().substring(0, newp) + "<br/>";


                                        if ($('#labeltotal6').text() == "excepções?") {
                                            
                                            var newc = $('#labellinha' + ll).text().search("c)");
                                            newstring += $('#labellinha' + ll).text().substring(newc, end) + "<br/>";
                                      
                                        }


                                        if ($('#labeltotal7').text() == ">TOP5?") {
                                        
                                            var newt = $('#labellinha' + ll).text().search(">top5");
                                            newstring += $('#labellinha' + ll).text().substring(newt, end);
                                          
                                        }
                                     
                                    }
                               
                                    pre2port();
                                   
                                }
                              
                                showchange(ll);
                               
                                mudado = organismo + "" + $('#pvp' + ll).children('option:selected').index() + "" + $('#port' + ll).children('option:selected').index() + "" + ll;
                           
                                if (organismo != "45" && organismo != "49") {

                                    //if ($('#labeltotal5').text() == "despachos?") {

                                    if ($('#labeltotal5').css("visibility") == "hidden") {
                                        $('#labeltotal5').css("visibility", "visible");

                                    }


                                    if ($('#port' + ll).children('option').length > 1) {
                                        $('#labeltotal5').html("despachos?");
                                        var sel2rpre = document.getElementById("pr" + ll);
                                        var rdp1 = sel2rpre.options[3].value;
                                        var rdp2 = sel2rpre.options[4].value;
                                        var rdp3 = sel2rpre.options[5].value;
                                        var rdp4 = sel2rpre.options[6].value;
                                        var rdp5 = sel2rpre.options[7].value;
                                        var rdp6 = sel2rpre.options[8].value;
                                        var rdp7 = sel2rpre.options[9].value;
                                        var rdp8 = sel2rpre.options[10].value;

                                        choose2port(ll, rdp1, rdp2, rdp3, rdp4, rdp5, rdp6, rdp7, rdp8);
                                    }



                                    if ($('#labeltotal0').text() == '') { //sem alineas //tornar yellow se sem alineas
                                     
                                        $('#semaforo').attr("src", "/img/yellow.gif");
                                  
                                    }

                                    //} //fecha labeltotal5
                                } else { }

                                if (organismo == "45" || organismo == "49") {
                                    if ($('#port' + ll).css("visibility") == "hidden") {

                                        $('#port' + ll).css("visibility", "visible");

                                    }
                                    if ($('#labeltotal5').text() == "despachos?") {

                                        if ($('#labeltotal5').css("visibility") == "visible") {

                                            $('#labeltotal5').css("visibility", "hidden");
                                            $('#labellinha' + ll).html(newstring);

                                        }
                                    } //fecha labeltotal5

                                    if ($('#labeltotal0').text() == '') { //sem alineas 

                                        if (($('#labeltotal6').text() != "excepções?") || ($('#labeltotal7').text() != ">TOP5?")) { //tornar green se sem alineas nem avisos
                                      
                                            $('#semaforo').attr("src", "/img/green.gif");
                                        
                                        } //fecha if excepçoes etc
                                    }
                                }

                                if (organismo == "DS") {
                                    //ds
                                }
                            } //fecha texto de labellinha+ll vazio ou não
                        }//fecha verif mudado
                    } //fecha for
          
            } else { //não mostrado

                if ($('#p1').val() == "") {
                    $("#p1").focus();
                } //fecha p1.text vazio
            } //fecha mostrado ou não
            organterior = organismo;
        }

    }); //fecha função radio change
  
  
    $("#btnSubmit").click(function () {
       
        if (document.getElementById("btnSubmit").style.visibility != "hidden") {
            
            toggle();
            var organismo = jQuery('input[name=radios]:checked').val();
            //first get all parameters if any
         
            var param1 = $('#p1').val(),
                  param2 = $('#p2').val(),
                  param3 = $('#p3').val(),
                  param4 = $('#p4').val(),
                  param5 = $('#a1').val(),
                  param6 = $('#a2').val(),
                  param7 = $('#a3').val(),
                  param8 = $('#a4').val();


            $.ajax({

                type: "POST",
                //dataType:  "text/html",
                contentType: "application/json; charset=utf-8",

                url: "../prescavi/comparar",


                data: "{presc1:'" + param1 + "', aviad1:'" + param5 + "',presc2:'" + param2 + "', aviad2:'" + param6 + "', presc3:'" + param3 + "', aviad3:'" + param7 + "', presc4:'" + param4 + "', aviad4:'" + param8 + "'}",




                success: function (blue) {
                    //$(document).html(blue); //swapnil commented

                    //swapnil added
                    debugger
             
                    if (blue != null) {

                        if (blue.length > 0) {

                            var labeltotal = (blue[0][1]);
                            if ((blue[0][1]) == "l") {
                                labeltotal = (blue[0][1]).toUpperCase();
                            }


                            $('#labeltotal0').html(labeltotal);
                            tooltipalinea("labeltotal0", (blue[0][1]));

                            //$('#labeltotal1').html(blue[0][5] + ' ' + blue[0][1] + ' ' + blue[0][2] + ' ' + blue[0][3]);
                            //$('#labeltotal2').html(blue[1][0] + ' ' + blue[1][1] + ' ' + blue[1][2] + ' ' + blue[1][3]);
                            //$('#labeltotal3').html(blue[1][0] + ' ' + blue[1][1] + ' ' + blue[1][2] + ' ' + blue[1][3]);
                            //$('#labeltotal4').html(blue[1][0] + ' ' + blue[1][1] + ' ' + blue[1][2] + ' ' + blue[1][3]);
                            //$('#labeltotal5').html(blue[1][0] + ' ' + blue[1][1] + ' ' + blue[1][2] + ' ' + blue[1][3]);
                            //$('#labeltotal6').html(blue[1][0] + ' ' + blue[1][1] + ' ' + blue[1][2] + ' ' + blue[1][3]);

                            //for (var i = 0; i < blue.length; i++) {
                            //    var obj = blue[i];

                            //    if (obj[0].length > 0) {
                            //        if (obj[1] != "a") {
                            //            $("#labellinha"+i).css("fontSize", "20px");
                            //            $("#labellinha"+i).css("fontWeight", "bold");
                            //            $("#labellinha"+i).css("color", (obj[5]));
                            //            $("#labellinha"+i).css("text-shadow", "#000 1px 1px 1px");
                            //            $("#labellinha"+i).css("-webkit - font - smoothing", "antialiased");
                            //            $("#labeltotal"+i).html(obj[0]);
                            //            var qual+i+p = (obj[0]).substring(0, 1);
                            //            $("#p" + qual1p).css("background-color", (obj[5]));
                            //            $("#a"+i).css("background-color", (obj[5]));
                            //            tooltipalinea("labellinha"+i, (obj[1]));
                            //            tooltipalinea("labeltotal"+i, (obj[1]));
                            //            if (obj[1] == "w") {
                            //                $("#showRight").click();
                            //                $("#cnpemquery").text = $("#a"+i).val;
                            //                querycnpem();
                            //            }
                            //        }
                            //    }
                            //}

                          
                            if (blue[1][0].length > 0) {
                           
                                var qual1p = (blue[1][6]);
                                if ((blue[1][2] == true && organismo != "45" && organismo != "49") || blue[1][3] == true || blue[1][4] == true) { //para dar amarelo e grosso em linhas com excep top5 e ports
                                
                                    $("#labellinha1").css("fontSize", "20px");
                                    $("#labellinha1").css("fontWeight", "bold");
                                    $("#labellinha1").css("color", "yellow");
                                    $("#labellinha1").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha1").css("-webkit - font - smoothing", "antialiased");
                                }
                                $('#labellinha1').html(blue[1][0]);
                                if (blue[1][1] != "a") { //para dar cor do erro e grosso em linhas com erro e acrescentar Às linhas totais
                                  
                                    $("#labellinha1").css("fontSize", "20px");
                             
                                    $("#labellinha1").css("fontWeight", "bold");
                               
                                    $("#labellinha1").css("color", (blue[1][5]));//porque não red?
                                
                                    $("#labellinha1").css("text-shadow", "#000 1px 1px 1px");
                         
                                    $("#labellinha1").css("-webkit - font - smoothing", "antialiased");
                             
                                 
                                    $("#labeltotal1").html(blue[1][0]);
                           
                                    //próximas 3 dão cor aos inputs p e a do cruzamento com erro
                                    $("#p" + qual1p).css("background-color", (blue[1][5]));
                                
                                    $("#a1").css("background-color", (blue[1][5]));
                                
                                    tooltipalinea("labellinha1", (blue[1][1])); //próximas 2 fazem tooltip de descrição da alínea por linha e linhatotal
                               
                                    tooltipalinea("labeltotal1", (blue[1][1]));
                                

                                    if (blue[1][1] == "w") { //enviar para cnpemfill     o mal é que faz duas vezes (abre e fecha ou fecha e abre) e ainda não resolvi como ver se está aberto, etc
                                        if (classie.has(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open')) {//se estiver aberto
                                            //deixa aberto
                                        } else {//se estiver fechado, abre
                                            classie.toggle(document.getElementById('showRight'), 'active');
                                            classie.toggle(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open');
                                        }
                                        $("#cnpemquery").val(blue[qual1p+4][0]);
                                        var cr = jQuery.Event("keypress");
                                        cr.which = 13;
                                        $("#cnpemquery").trigger(cr);
                                    }


                                }
                          
                                if (blue[1][1] == "a" || blue[1][1] == "L" || blue[1][1] == "h") {
                                   
                                    choose2port(1, blue[9][20], blue[9][25], blue[9][26], blue[9][27], blue[9][21], blue[9][23], blue[9][24], blue[9][22]);
                             
                                    choose2pr(1, blue[9][9], blue[9][6], blue[9][12], blue[9][20], blue[9][25], blue[9][26], blue[9][27], blue[9][21], blue[9][23], blue[9][24], blue[9][22]);
                                 
                                    choose2array(1, blue[9][8], blue[9][13], blue[9][14], blue[9][17], blue[9][18], blue[9][19]);
                                  
                                    //if (pvp1.options.length == 0) {
                                    show(1);
                                  
                                    //}

                                }

                            }

                            if (blue[2][0] != null) {
                                var qual2p = (blue[2][6]);
                                if ((blue[2][2] == true && organismo != "45" && organismo != "49") || blue[2][3] == true || blue[2][4] == true) {

                                    $("#labellinha2").css("fontSize", "20px");
                                    $("#labellinha2").css("fontWeight", "bold");
                                    $("#labellinha2").css("color", "yellow");
                                    $("#labellinha2").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha2").css("-webkit - font - smoothing", "antialiased");
                                }
                                $('#labellinha2').html(blue[2][0]);
                                if (blue[2][1] != "a") {
                                    $("#labellinha2").css("fontSize", "20px");
                                    $("#labellinha2").css("fontWeight", "bold");
                                    $("#labellinha2").css("color", "red");
                                    $("#labellinha2").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha2").css("-webkit - font - smoothing", "antialiased");
                                    $("#labeltotal2").html(blue[2][0]);
                                   
                                    $("#p" + qual2p).css("background-color", (blue[2][5]));
                                    $("#a2").css("background-color", (blue[2][5]));
                                    tooltipalinea("labellinha2", (blue[2][1]));
                                    tooltipalinea("labeltotal2", (blue[2][1]));

                                    if (blue[2][1] == "w") { //enviar para cnpemfill     o mal é que faz duas vezes (abre e fecha ou fecha e abre) e ainda não resolvi como ver se está aberto, etc
                                        if (classie.has(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open')) {//se estiver aberto
                                            //deixa aberto
                                        } else {//se estiver fechado, abre
                                            classie.toggle(document.getElementById('showRight'), 'active');
                                            classie.toggle(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open');
                                        }
                                      
                                        $("#cnpemquery").val(blue[qual2p+4][0]);
                                        var cr = jQuery.Event("keypress");
                                        cr.which = 13;
                                        $("#cnpemquery").trigger(cr);
                                    }
                                }
                   
                                if (blue[2][1] == "a" || blue[2][1] == "L" || blue[2][1] == "h") {
                             
                                    
                              
                                    choose2port(2, blue[10][20], blue[10][25], blue[10][26], blue[10][27], blue[10][21], blue[10][23], blue[10][24], blue[10][22]);
                                
                                    choose2pr(2, blue[10][9], blue[10][6], blue[10][12], blue[10][20], blue[10][25], blue[10][26], blue[10][27], blue[10][21], blue[10][23], blue[10][24], blue[10][22]);
                                
                                    choose2array(2, blue[10][8], blue[10][13], blue[10][14], blue[10][17], blue[10][18], blue[10][19]);
                                
                                    //if (pvp2.options.length == 0) {
                                    show(2);
                          
                                    //}
                                }
                            }

                            if (blue[3][0] != null) {
                                var qual3p = (blue[3][6]);
                                if ((blue[3][2] == true && organismo != "45" && organismo != "49") || blue[3][3] == true || blue[3][4] == true) {
                                    $("#labellinha3").css("fontSize", "20px");
                                    $("#labellinha3").css("fontWeight", "bold");
                                    $("#labellinha3").css("color", "yellow");
                                    $("#labellinha3").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha3").css("-webkit - font - smoothing", "antialiased");
                                }
                                $('#labellinha3').html(blue[3][0]);
                                if (blue[3][1] != "a") {
                                    $("#labellinha3").css("fontSize", "20px");
                                    $("#labellinha3").css("fontWeight", "bold");
                                    $("#labellinha3").css("color", "red");
                                    $("#labellinha3").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha3").css("-webkit - font - smoothing", "antialiased");
                                    $("#labeltotal3").html(blue[3][0]);
                            
                                    $("#p" + qual3p).css("background-color", (blue[3][5]));
                                    $("#a3").css("background-color", (blue[3][5]));
                                    tooltipalinea("labellinha3", (blue[3][1]));
                                    tooltipalinea("labeltotal3", (blue[3][1]));

                                    if (blue[3][1] == "w") { //enviar para cnpemfill     o mal é que faz duas vezes (abre e fecha ou fecha e abre) e ainda não resolvi como ver se está aberto, etc
                                        if (classie.has(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open')) {//se estiver aberto
                                            //deixa aberto
                                        } else {//se estiver fechado, abre
                                            classie.toggle(document.getElementById('showRight'), 'active');
                                            classie.toggle(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open');
                                        }
                                        $("#cnpemquery").val(blue[qual3p+4][0]);
                                        var cr = jQuery.Event("keypress");
                                        cr.which = 13;
                                        $("#cnpemquery").trigger(cr);

                                    }

                                }
                                if (blue[3][1] == "a" || blue[3][1] == "L" || blue[3][1] == "h") {
                                    
                                    choose2port(3, blue[11][20], blue[11][25], blue[11][26], blue[11][27], blue[11][21], blue[11][23], blue[11][24], blue[11][22]);
                                    choose2pr(3, blue[11][9], blue[11][6], blue[11][12], blue[11][20], blue[11][25], blue[11][26], blue[11][27], blue[11][21], blue[11][23], blue[11][24], blue[11][22]);
                                    choose2array(3, blue[11][8], blue[11][13], blue[11][14], blue[11][17], blue[11][18], blue[11][19]);

                                    //if (pvp3.options.length == 0) {
                                    show(3);

                                    //}
                                }
                            }

                            if (blue[4][0] != null) {
                                var qual4p = (blue[4][6]);

                                if ((blue[4][2] == true && organismo != "45" && organismo != "49") || blue[4][3] == true || blue[4][4] == true) {

                                    $("#labellinha4").css("fontSize", "20px");
                                    $("#labellinha4").css("fontWeight", "bold");
                                    $("#labellinha4").css("color", "yellow");
                                    $("#labellinha4").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha4").css("-webkit - font - smoothing", "antialiased");
                                }
                                $('#labellinha4').html(blue[4][0]);
                                if (blue[4][1] != "a") {

                                    $("#labellinha4").css("fontSize", "20px");
                                    $("#labellinha4").css("fontWeight", "bold");
                                    $("#labellinha4").css("color", "red");
                                    $("#labellinha4").css("text-shadow", "#000 1px 1px 1px");
                                    $("#labellinha4").css("-webkit - font - smoothing", "antialiased");
                                    $("#labeltotal4").html(blue[4][0]);
                               
                                    $("#p" + qual4p).css("background-color", (blue[4][5]));
                                    $("#a4").css("background-color", (blue[4][5]));
                                    tooltipalinea("labellinha4", (blue[4][1]));
                                    tooltipalinea("labeltotal4", (blue[4][1]));

                                    if (blue[4][1] == "w") { //enviar para cnpemfill     o mal é que faz duas vezes (abre e fecha ou fecha e abre) e ainda não resolvi como ver se está aberto, etc
                                        if (classie.has(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open')) {//se estiver aberto
                                            //deixa aberto
                                        } else {//se estiver fechado, abre
                                            classie.toggle(document.getElementById('showRight'), 'active');
                                            classie.toggle(document.getElementById('cbp-spmenu-s2'), 'cbp-spmenu-open');
                                        }
                                        $("#cnpemquery").val(blue[qual4p+4][0]);
                                        var cr = jQuery.Event("keypress");
                                        cr.which = 13;
                                        $("#cnpemquery").trigger(cr);
                                    }
                                }
                                if (blue[4][1] == "a" || blue[4][1] == "L" || blue[4][1] == "h") {
                                   
                                    choose2port(4, blue[12][20], blue[12][25], blue[12][26], blue[12][27], blue[12][21], blue[12][23], blue[12][24], blue[12][22]);
                                    choose2pr(4, blue[12][9], blue[12][6], blue[12][12], blue[12][20], blue[12][25], blue[12][26], blue[12][27], blue[12][21], blue[12][23], blue[12][24], blue[12][22]);
                                    choose2array(4, blue[12][8], blue[12][13], blue[12][14], blue[12][17], blue[12][18], blue[12][19]);
                                    //if (pvp4.options.length == 0) {
                                    show(4);
                                    //}
                                }
                            }

                            //pelo blue[0] nãodeu
                            if ((blue[1][2] == true || blue[2][2] == true || blue[3][2] == true || blue[4][2]) && (organismo == "01" || organismo == "48")) {
                                //aviados códigos que podem ter portarias
                                //if (organismo != "45" || organismo != "49"){
                                $("#labeltotal5").text('despachos?');
                                //}
                            }

                            if (blue[1][3] == true || blue[2][3] == true || blue[3][3] == true || blue[4][3]) {
                                //excepções avisadas em aviados
                                $("#labeltotal6").text('excepções?');
                            }

                            if (blue[1][4] == true || blue[2][4] == true || blue[3][4] == true || blue[4][4]) {
                                //top5 avisado em aviados
                                $("#labeltotal7").text('>TOP5?');
                            }





                            var corcor = blue[0][5];

                            //se semaforo é verde muda para amarelo se algum dos 4 tiver portaria
                            if ((blue[1][2] == true || blue[2][2] == true || blue[3][2] == true || blue[4][2] == true) && (organismo == "01" || organismo == "48")) {

                                if (corcor == "green") {
                                    corcor = "yellow";
                                }
                            }
                            mudarimagem(corcor);

                            show(0);

                            if (blue[12][0] > 0) {
                                $('#a4').prop('title', 'cnpem: ' + blue[12][15] + '\r\n' + 'nome: ' + blue[12][2] + '\r\n' + 'dci: ' + blue[12][1] + '\r\n' + 'forma: ' + blue[12][3] + '\r\n' + 'dosagem: ' + blue[12][4] + '\r\n' + 'x ' + blue[12][5]);
                            }
                            if (blue[4][3] == true) {//se excep
                                if ((100 * blue[12][8]) > (100 * blue[parseFloat(qual4p) + 4][8])) {//pvpa>pvpp
                                    $('#a4').prop('title', $('#a4').prop('title') + '\r\n' + 'pvp (€' + blue[12][8] + ') mais caro que prescrito(€' + blue[parseFloat(qual4p) + 4][8] + ')! <= excepção');
                                } else {//pvpa!>pvpp
                                    $('#a4').prop('title', $('#a4').prop('title') + '\r\n' + 'pvp (€' + blue[12][8] + ') menor que prescrito => opção <= excepção');
                                }
                            }
                            if (blue[4][4] == true) {//se top
                                $('#a4').prop('title', $('#a4').prop('title') + '\r\n' + 'pvp (€' + blue[12][8] + ') superior 5º mais barato (€' + blue[12][12] + ') => opção');//+ '\r\n' + 'lab= ' + blue[12][11] + '\r\n' + 'gen= ' + blue[12][10] + '\r\n' + 'por dci= ' + blue[12][16]);
                            }

                            if (blue[11][0] > 0) {
                                $('#a3').prop('title', 'cnpem: ' + blue[11][15] + '\r\n' + 'nome: ' + blue[11][2] + '\r\n' + 'dci: ' + blue[11][1] + '\r\n' + 'forma: ' + blue[11][3] + '\r\n' + 'dosagem: ' + blue[11][4] + '\r\n' + 'x ' + blue[11][5]);
                            }
                            if (blue[3][3] == true) {//se excep
                                if ((100 * blue[11][8]) > (100 * blue[parseFloat(qual3p) + 4][8])) {//pvpa>pvpp
                                    $('#a3').prop('title', $('#a3').prop('title') + '\r\n' + 'pvp (€' + blue[11][8] + ') mais caro que prescrito(€' + blue[parseFloat(qual3p) + 4][8] + ')! <= excepção');
                                } else {//pvpa!>pvpp
                                    $('#a3').prop('title', $('#a3').prop('title') + '\r\n' + 'pvp (€' + blue[11][8] + ') menor que prescrito => opção <= excepção');
                                }
                            }
                            if (blue[3][4] == true) {//se top
                                $('#a3').prop('title', $('#a3').prop('title') + '\r\n' + 'pvp (€' + blue[11][8] + ') superior 5º mais barato (€' + blue[11][12] + ') => opção');//+ '\r\n' + 'lab= ' + blue[11][11] + '\r\n' + 'gen= ' + blue[11][10] + '\r\n' + 'por dci= ' + blue[11][16]);
                            }

                            if (blue[10][0] > 0) {
                                $('#a2').prop('title', 'cnpem: ' + blue[10][15] + '\r\n' + 'nome: ' + blue[10][2] + '\r\n' + 'dci: ' + blue[10][1] + '\r\n' + 'forma: ' + blue[10][3] + '\r\n' + 'dosagem: ' + blue[10][4] + '\r\n' + 'x ' + blue[10][5]);
                            }

                            if (blue[2][3] == true) {


                                if ((100 * blue[10][8]) > (100 * blue[parseFloat(qual2p) + 4][8])) {

                                    $('#a2').prop('title', $('#a2').prop('title') + '\r\n' + 'pvp (€' + blue[10][8] + ') mais caro que prescrito(€' + blue[parseFloat(qual2p) + 4][8] + ')! <= excepção');
                                } else {

                                    $('#a2').prop('title', $('#a2').prop('title') + '\r\n' + 'pvp (€' + blue[10][8] + ') menor que prescrito => opção <= excepção');
                                }

                            }
                            if (blue[2][4] == true) {

                                $('#a2').prop('title', $('#a2').prop('title') + '\r\n' + 'pvp (€' + blue[10][8] + ') superior 5º mais barato (€' + blue[10][12] + ') => opção');//+ '\r\n' + 'lab= ' + blue[10][11] + '\r\n' + 'gen= ' + blue[10][10] + '\r\n' + 'por dci= ' + blue[10][16]);
                            }

                            if (blue[9][0] > 0) {
                                $('#a1').prop('title', 'cnpem: ' + blue[9][15] + '\r\n' + 'nome: ' + blue[9][2] + '\r\n' + 'dci: ' + blue[9][1] + '\r\n' + 'forma: ' + blue[9][3] + '\r\n' + 'dosagem: ' + blue[9][4] + '\r\n' + 'x ' + blue[9][5]);
                            }
                            if (blue[1][3] == true) {
                                if ((100 * blue[9][8]) > (100 * blue[parseFloat(qual1p) + 4][8])) {
                                    $('#a1').prop('title', $('#a1').prop('title') + '\r\n' + 'pvp (€' + blue[9][8] + ') mais caro que prescrito(€' + blue[parseFloat(qual1p) + 4][8] + ')! <= excepção');
                                } else {
                                    $('#a1').prop('title', $('#a1').prop('title') + '\r\n' + 'pvp (€' + blue[9][8] + ') menor que prescrito => opção <= excepção');
                                }
                            }
                            if (blue[1][4] == true) {
                                $('#a1').prop('title', $('#a1').prop('title') + '\r\n' + 'pvp (€' + blue[9][8] + ') superior 5º mais barato (€' + blue[9][12] + ') => opção');//+ '\r\n' + 'lab= ' + blue[9][11] + '\r\n' + 'gen= ' + blue[9][10] + '\r\n' + 'por dci= ' + blue[9][16]);
                            }






                            if (blue[8][0] > 0) {
                                $('#p4').prop('title', 'cnpem: ' + blue[8][15] + '\r\n' + 'nome: ' + blue[8][2] + '\r\n' + 'dci: ' + blue[8][1] + '\r\n' + 'forma: ' + blue[8][3] + '\r\n' + 'dosagem: ' + blue[8][4] + '\r\n' + 'x ' + blue[8][5] + '\r\n' + 'pvp= €' + blue[8][8] + '\r\n' + 'top5');
                            }
                            if (100 * blue[8][12] > 0) {
                                $('#p4').prop('title', $('#p4').prop('title') + '= €' + blue[8][12]);
                            }

                            if (blue[7][0] > 0) {
                                $('#p3').prop('title', 'cnpem: ' + blue[7][15] + '\r\n' + 'nome: ' + blue[7][2] + '\r\n' + 'dci: ' + blue[7][1] + '\r\n' + 'forma: ' + blue[7][3] + '\r\n' + 'dosagem: ' + blue[7][4] + '\r\n' + 'x ' + blue[7][5] + '\r\n' + 'pvp= €' + blue[7][8] + '\r\n' + 'top5');
                            }
                            if (100 * blue[7][12] > 0) {
                                $('#p3').prop('title', $('#p3').prop('title') + '= €' + blue[7][12]);
                            }

                            if (blue[6][0] > 0) {
                                $('#p2').prop('title', 'cnpem: ' + blue[6][15] + '\r\n' + 'nome: ' + blue[6][2] + '\r\n' + 'dci: ' + blue[6][1] + '\r\n' + 'forma: ' + blue[6][3] + '\r\n' + 'dosagem: ' + blue[6][4] + '\r\n' + 'x ' + blue[6][5] + '\r\n' + 'pvp= €' + blue[6][8] + '\r\n' + 'top5');
                            }
                            if (100 * blue[6][12] > 0) {
                                $('#p2').prop('title', $('#p2').prop('title') + '= €' + blue[6][12]);
                            }

                            if (blue[5][0] > 0) {
                                $('#p1').prop('title', 'cnpem: ' + blue[5][15] + '\r\n' + 'nome: ' + blue[5][2] + '\r\n' + 'dci: ' + blue[5][1] + '\r\n' + 'forma: ' + blue[5][3] + '\r\n' + 'dosagem: ' + blue[5][4] + '\r\n' + 'x' + blue[5][5] + '\r\n' + 'pvp= €' + blue[5][8] + '\r\n' + 'top5');
                            }
                            if (100 * blue[5][12] > 0) {
                                $('#p1').prop('title', $('#p1').prop('title') + '= €' + blue[5][12]);
                            }

                        }//fim blue.length
                    }//fim blue != null

                },




                //  call on ajax call failure
                error: function (xhr, textStatus, error) {

                    //called on ajax call success
                    console.log("readyState: " + xhr.readyState);
                    console.log("responseText: " + xhr.responseText);
                    console.log("status: " + xhr.status);
                    console.log("textStatus: " + textStatus);
                    alert("Error: " + error);
                }
            });

        } else {
          
            //do nothing
        }
    });
   
});


function mudarimagem(colour) {

    var image = document.getElementById('semaforo');
    image.src = "../img/" + colour + ".gif"

    document.getElementById("btnlimpar").style.visibility = "visible"
    //document.getElementById("btnlimpar").focus()  //não quero que 2enters façam o comparar e logo a seguir limpem

};







function toggle() {
    document.getElementById("btnSubmit").style.visibility = "hidden"

}


$(window).keypress(function (e) {

    if (e.which == 32) {
        var focado5 = document.getElementById(document.activeElement.id);
        var classefocada = document.getElementsByClassName(document.activeElement.className);
    


        if (focado5.id == "cnpemquery") {
             document.getElementById(focado5.id).value = "";
            $('.cnpem').each(function (index, data) {
                $(this).html('');
            });
        } else {
            document.getElementById("btnlimpar").click();
        }
    }
});

$("#btnlimpar").keydown(function (e) {

    if (e.which == 13) {
        e.preventDefault;
    }
});



function limpar() {
    document.getElementById("semaforo").src = "../img/empty.gif";
    document.getElementById("labellinha1").innerHTML = "";
    document.getElementById("labellinha2").innerHTML = "";
    document.getElementById("labellinha3").innerHTML = "";
    document.getElementById("labellinha4").innerHTML = "";
    document.getElementById("labellinha1").style.fontSize = "";
    document.getElementById("labellinha2").style.fontSize = "";
    document.getElementById("labellinha3").style.fontSize = "";
    document.getElementById("labellinha4").style.fontSize = "";
    document.getElementById("labellinha1").style.fontWeight = "normal";
    document.getElementById("labellinha2").style.fontWeight = "normal";
    document.getElementById("labellinha3").style.fontWeight = "normal";
    document.getElementById("labellinha4").style.fontWeight = "normal";
    document.getElementById("labellinha1").style.color = "black";
    document.getElementById("labellinha2").style.color = "black";
    document.getElementById("labellinha3").style.color = "black";
    document.getElementById("labellinha4").style.color = "black";
    document.getElementById("btnlimpar").style.visibility = "hidden";
    document.getElementById("p1").focus();
      document.getElementById("labellinha1").setAttribute('title', "");
    document.getElementById("labellinha2").setAttribute('title', "");
    document.getElementById("labellinha3").setAttribute('title', "");
    document.getElementById("labellinha4").setAttribute('title', "");
    document.getElementById("p1").setAttribute('title', "");
    document.getElementById("p2").setAttribute('title', "");
    document.getElementById("p3").setAttribute('title', "");
    document.getElementById("p4").setAttribute('title', "");
    document.getElementById("a1").setAttribute('title', "");
    document.getElementById("a2").setAttribute('title', "");
    document.getElementById("a3").setAttribute('title', "");
    document.getElementById("a4").setAttribute('title', "");
    document.getElementById("resetbtn").click();
    document.getElementById("labeltotal0").setAttribute('title', "");
    document.getElementById("labeltotal0").innerHTML = "";
    document.getElementById("labeltotal1").innerHTML = "";
    document.getElementById("labeltotal2").innerHTML = "";
    document.getElementById("labeltotal3").innerHTML = "";
    document.getElementById("labeltotal4").innerHTML = "";
    document.getElementById("labeltotal5").innerHTML = "";
    document.getElementById("labeltotal6").innerHTML = "";
    document.getElementById("labeltotal7").innerHTML = "";
    document.getElementById("p1").style.backgroundColor = "lightblue";
    document.getElementById("p2").style.backgroundColor = "lightblue";
    document.getElementById("p3").style.backgroundColor = "lightblue";
    document.getElementById("p4").style.backgroundColor = "lightblue";
    document.getElementById("a1").style.backgroundColor = "lightblue";
    document.getElementById("a2").style.backgroundColor = "lightblue";
    document.getElementById("a3").style.backgroundColor = "lightblue";
    document.getElementById("a4").style.backgroundColor = "lightblue";
    document.getElementById("pvp1").innerHTML = "";
    document.getElementById("pvp2").innerHTML = "";
    document.getElementById("pvp3").innerHTML = "";
    document.getElementById("pvp4").innerHTML = "";
    document.getElementById("sns1").innerHTML = "";
    document.getElementById("sns2").innerHTML = "";
    document.getElementById("sns3").innerHTML = "";
    document.getElementById("sns4").innerHTML = "";
    
    document.getElementById("pvp1").style.visibility = "hidden";
    document.getElementById("pvp2").style.visibility = "hidden";
    document.getElementById("pvp3").style.visibility = "hidden";
    document.getElementById("pvp4").style.visibility = "hidden";
    document.getElementById("port1").style.visibility = "hidden";
    document.getElementById("port2").style.visibility = "hidden";
    document.getElementById("port3").style.visibility = "hidden";
    document.getElementById("port4").style.visibility = "hidden";
    document.getElementById("snstotal").style.visibility = "hidden";
    document.getElementById("snstotal").innerHTML = "";
    
    mudado = "00000";



    for (ss = 10; ss >= 0; ss--) {
      
        for (rr = 1; rr < 5; rr++) {
            if (document.getElementById("pr"+rr).length >= ss){
                document.getElementById("pr" + rr).remove(ss);
           }
            if (document.getElementById("pvp"+rr).length >=ss){
                document.getElementById("pvp" + rr).remove(ss);
            }
            if (document.getElementById("port" + rr).length >=ss) {
                document.getElementById("port" + rr).remove(ss);
            }
        }
    }
};







$(document).ready(function () {
    $(cnpemquery).keypress(function (e) {
        if (e.which == 13) {

            $('.cnpem').each(function (index, data) {
                $(this).html('');
            });

            $('#labelcnpem1').html("aguarde");
            var param = $('#cnpemquery').val();

            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: "../prescavi/cnpemfill",
                data: "{pedido:'" + param + "'}",

                success: function (dev) {
                  
                    debugger
                    if (dev != null) {
                        if (dev.length > 0) {
                            for (var i = 1; dev.length; i++) {
                                var linhacnpem = '' + i;
                                linhacnpem = '#labelcnpem' + linhacnpem;
                                var index = parseFloat(i) - 1;
                                if (dev[index].length > 0) {
                                    $(linhacnpem).html(dev[index][0]);
                                    $(linhacnpem).prop('title', 'nome: ' + dev[index][1] + '\r\n' + 'dci: ' + dev[index][2] + '\r\n' + 'forma: ' + dev[index][3] + '\r\n' + 'dosagem: ' + dev[index][4] + '\r\n' + 'x ' + dev[index][5] + '\r\n' + 'lab= ' + dev[index][6]);
                                } else {
                                    if (index == 0) {
                                        $('#labelcnpem1').html("não foram encontrados códigos que satisfaçam este cnpem");
                                    }
                                }
                            }
                        } else {
                            $('#labelcnpem1').html("não foram encontrados códigos que satisfaçam este cnpem");
                        } //fimdevlength
                    } //fim dev != null
                },

                //  call on ajax call failure
                error: function (xhr, textStatus, error) {
                 
                    //called on ajax call success
                    console.log("readyState: " + xhr.readyState);
                    console.log("responseText: " + xhr.responseText);
                    console.log("status: " + xhr.status);
                    console.log("textStatus: " + textStatus);
                    alert("Error: " + error);
                }
            }); //fecha ajax
        } //fecha if
    });
});





function choose2array(wch, now, last, b4last, _2b4last, _3b4last, _4b4last) {
 
    var hoje2 = new Date();
    var mes2 = hoje2.getMonth() + 1;
    var arrayold = [now, last, b4last, _2b4last, _3b4last, _4b4last]
    
    var arraypvp = [];
    
    switch (mes2) {
        case 1:
        case 4:
        case 7:
        case 10:
            var mostrar = 4;
            break;
        case 2:
        case 5:
        case 8:
        case 11:
            var mostrar = 5;
            break;
        case 3:
        case 6:
        case 9:
        case 12:
            var mostrar = 6;
            break;
       
    }

    for (var m = 1; m <= mostrar; m++) {
       
        //var found = jQuery.inArray(arrayold[m], arraypvp);
        if (jQuery.inArray(arrayold[m-1], arraypvp) < 0) {   // Element was not found, add it.
            
            if (arrayold[m-1] != 0) {
                arraypvp.push(arrayold[m-1]);
                
            }
        }

    }
 
    var sel = document.getElementById('pvp' + wch);
    if (sel.options.length == 0) {
        if (arraypvp.length > 0) {
            for (var i = 0; i < arraypvp.length; i++) {
                var opt = document.createElement('option');
                opt.innerHTML = arraypvp[i];
                opt.value = arraypvp[i];
                sel.appendChild(opt);
            }

            //  document.getElementById("pvp"+wch).style.visibility = "visible";   //semaforo
        } else {
            var opt = document.createElement('option');
            opt.innerHTML = 0;
            opt.value = 0;
            sel.appendChild(opt);
        }
    }
    
    
   
}


function choose2pr(w, r, c, t, pd1, pd2, pd3, pd4, pd5, pd6, pd7, pd8) {
   
    var arraypr = [];
  
    if (arraypr.length == 0) {
        arraypr.push(r);
        arraypr.push(c);
        arraypr.push(t);
        
        pd1 ? arraypr.push(1) : arraypr.push(-1);
        pd2 ? arraypr.push(1) : arraypr.push(-1);
        pd3 ? arraypr.push(1) : arraypr.push(-1);
        pd4 ? arraypr.push(1) : arraypr.push(-1);
        pd5 ? arraypr.push(1) : arraypr.push(-1);
        pd6 ? arraypr.push(1) : arraypr.push(-1);
        pd7 ? arraypr.push(1) : arraypr.push(-1);
        pd8 ? arraypr.push(1) : arraypr.push(-1);
        
         
    }
   
    var sel2 = document.getElementById('pr' + w);
   
    if (sel2.options.length == 0) {
      
        if (arraypr.length > 0) {
          
            for (var zz = 0; zz < arraypr.length; zz++) {
              
                var opt2 = document.createElement('option');
                opt2.innerHTML = arraypr[zz];
                opt2.value = arraypr[zz];
                sel2.appendChild(opt2);
               
            }

        }

    }
}



function showchange(chg) {
    
    show(chg);
   
    show(0)
    
}

function show(qq) {

   
    if (qq == 0) {
  
        var snstotal = parseFloat(0);
     
        if (document.getElementById("labeltotal0").innerHTML == "" || document.getElementById("labeltotal0").innerHTML == "h") {
      
            for (qqq = 1; qqq < 5; qqq++) {
          
                if (document.getElementById("sns" + qqq).innerHTML != "") {
             
                    document.getElementById("pvp" + qqq).style.visibility = "visible";
                    document.getElementById("sns" + qqq).style.visibility = "visible";
               
                    snstotal = snstotal + parseFloat(document.getElementById("sns" + qqq).innerHTML);
                  
                }

            }

        }


        if (snstotal != 0) {
     
            snstotal = (snstotal).toFixed(2);
           
            snstotal = snstotal.replace(".", ",");
          
            snstotal = "€" + snstotal;
       
            document.getElementById("snstotal").style.visibility = "visible";
            document.getElementById("snstotal").innerHTML = snstotal;
         
        } else {
        
            document.getElementById("snstotal").innerHTML = "€0,00";
        }




    } else {
        
        document.getElementById("sns" + qq).innerHTML = calc(qq);
       
    }
}




function calc(w) {
 
    var o = document.querySelector('input[name = "radios"]:checked').value;
 
    var sele = document.getElementById("pvp" + w);
 
    var prtsel = document.getElementById("port" + w);
 
    var selpr = document.getElementById("pr" + w);
 
    var r = selpr.options[0].value;

    var c = selpr.options[1].value;
   
    var t = selpr.options[2].value;
  
    var p = sele.options[sele.selectedIndex].value;
    
    var prt = prtsel.options[prtsel.selectedIndex].value;

    var prt2 = prtsel.options.length;
  
    if (prt == "(sem)"  || (o != "45" && o != "49")) {//estva if ((prt == "(sem)" && prt2 == "1") || (o != "45" && o != "49")) {
        prt = "";
 
        if (prt == "(sem)" && prt2 == "1"){
            prtsel.style.visibility = "hidden";
    }
        switch (o) {
            case "01":
            case "45":
            case "46":
            case "DS":
                
                var novocomp = parseInt(c);
               
                break;
            case "48":
            case "49":
              
                if (c > 0) {
                
                    var novocomp = parseInt(c) + 15;
                    if (p <= t) {
                    
                        novocomp = 95;
                    }
                }
              
                if (novocomp > 95) {
                
                    novocomp = 95;
                }
                
                              
                if (c > novocomp) {
              
                    novocomp = c;
                }
               
                break;
            case "67":
            case "41":
               
                if (c > 0) {
                    var novocomp = 100;
                }
                break;
            case "42":
              
                var novocomp = 100;
                break;
        }
    } else {
       
        if (o == "49" || o == "45") {
        
            prtsel.style.visibility = "visible";
        } else {
      
            prtsel.style.visibility = "hidden";
        }
     
        switch (prt) {
            case "13020":
                novocomp = 37;
               
                break;
            case "10910":
            case "14123":
                novocomp = 69;
                break;
            case "lei6":
            case "1234":
            case "10279":
            case "10280":
            case "dores":
            case "5635":
            
                novocomp = 90;
                break;
            case "21094":
                novocomp = 100;
                break;
            case "medint":
                novocomp = 69;
                novocomp = 90;
            
        }
       
        novocomp = parseInt(novocomp);
        
        if (o == "49") {
            
            novocomp = novocomp + 15;
         
            if (novocomp > 95) {
                novocomp = 95;
            }
            if (p <= t) {

                novocomp = 95;
            }
            if (c > novocomp) {
                novocomp = c;
            }
        }
    }

    novocomp = novocomp / 100;
    
    if (novocomp > 0) {
       
        if ((r > 0) && (o != "42") && (o != "67") && (o != "41")) {
         
            var sns = (novocomp * r).toFixed(2);

        } else {
    
            var sns = (novocomp * p).toFixed(2);
        }

    } else {
   
        var sns = 0;
    }

    if (sns > p) {
        sns = p;
    }

   // alert("qual: " + w + "\n" + "pvp: " + p + "\n" + "pr: " + r + "\n" + "comp: " + c + "\n" + "top5: " + t);
    return sns;
    
}

function pre2port() {
    


   
        for (rr = 1; rr < 5; rr++) {
            for (ss = 0; ss < 11; ss++) {
                document.getElementById("port" + rr).remove(ss);
            }
            var mostradosns = document.getElementById("sns" + rr)
           
            if (mostradosns.innerHTML != "") {
              
                var sel2pre = document.getElementById("pr" + rr);
             
                var dp1 = sel2pre.options[3].value;
                var dp2 = sel2pre.options[4].value;
                var dp3 = sel2pre.options[5].value;
                var dp4 = sel2pre.options[6].value;
                var dp5 = sel2pre.options[7].value;
                var dp6 = sel2pre.options[8].value;
                var dp7 = sel2pre.options[9].value;
                var dp8 = sel2pre.options[10].value;
               
                choose2port(rr, dp1, dp2, dp3, dp4, dp5, dp6, dp7, dp8);
           
            


        }
    }
  
}



function choose2port(w, d1, d2, d3, d4, d5, d6, d7, d8) {
   
    var selpor = document.getElementById('port' + w);
    var op = document.querySelector('input[name = "radios"]:checked').value;
    
    var arrayport = [];
    var quallinha = document.getElementById("labellinha" + w);

    var br = quallinha.innerHTML.search("<br>");
   
    if (op == "45" || op == "49") {
        //selpor.style.visibility = "visible";
        
        if (arrayport.length == 0) {
            
            if (d1 == true) {
               
                arrayport.push("13020");
       
            }
            if (d2 == true) {
                arrayport.push("10910");
            }
            if (d3 == true) {
                arrayport.push("14123");
            }
            if (d4 == true) {
                arrayport.push("lei 6");
            }
            if (d5 == true) {
                arrayport.push("1234");
            }
            if (d6 == true) {
                arrayport.push("10279");
            }
            if (d7 == true) {
                arrayport.push("10280");
            }
            if (d8 == true) {
                arrayport.push("21094");
            }
            if (d1 != true && d2 != true && d3 != true && d4 != true && d5 != true && d6 != true && d7 != true && d8 != true) {
                selpor.style.visibility = "hidden";
            }
            //if (d9 == "true"){
            //    arrayport.push(" 5635");
            //}
            if (br >= 0) {
                var linhabr = quallinha.innerHTML.substring(0, br + 3);
                quallinha.innerHTML = linhabr;
                if (linhabr=="ok");{
                    document.getElementById("labellinha" + w).style.color = "black";
                }
            }
        }
    } else {
  
        var last22 = quallinha.innerHTML.slice(-22);

        if (d1 == true && last22 != "13020/2011, de 20/09 ?") {
            quallinha.innerHTML += "<br/> Despacho n.º 13020/2011, de 20/09 ?"
         
        }
        if (d2 == true && last22 != "10910/2011, de 22/04 ?") {
          
            quallinha.innerHTML += "<br/> Despacho n.º 10910/2011, de 22/04 ?"
        }
        if (d3 == true && (last22 != "14123/2009(2ª série), de 12/06 ?") && (last22 != " 1234/2006, de 29/12 ?")) {
          
            quallinha.innerHTML += "<br/> Despacho n.º 14123/2009(2ª série), de 12/06 ?"
            if (d5 == true) {
            
                quallinha.innerHTML += " e Despacho n.º 1234/2006, de 29/12 ?"
            }
        } else {
       
            if (d5 == true && last22 != " 1234/2006, de 29/12 ?") {
            
                quallinha.innerHTML += " <br/> Despacho n.º 1234/2006, de 29/12 ?"
            }
        }

        if (d4 == true && last22 != "n.º 6/2010, de 07/05 ?") {
            quallinha.innerHTML += "<br/> Lei n.º 6/2010, de 07/05 ?"
        }

        if (d6 == true && (last22 != "10279/2008, de 11/03 ?")&&(last22 != "10280/2008, de 11/03 ?")) {
            quallinha.innerHTML += "<br/> Despacho n.º 10279/2008, de 11/03 ?";
            if (d7 == true && last22 != "10280/2008, de 11/03 ?") {
                quallinha.innerHTML += " e Despacho n.º 10280/2008, de 11/03 ?";
            }
        } else {
            if (d7 == true && last22 != "10280/2008, de 11/03 ?") {
                quallinha.innerHTML += " <br/> Despacho n.º 10280/2008, de 11/03 ?";
            }
        }

        if (d8 == true && last22 != "º 21094/99, de 14/09 ?") {
            quallinha.innerHTML += "<br/> Despacho n.º 21094/99, de 14/09 ?"
        }

        selpor.style.visibility = "hidden";
    }
 
    arrayport.push("(sem)");
 
    if (selpor.options.length == 0) {
    
        if (arrayport.length > 0) {
     
            for (var pp = 0; pp < arrayport.length; pp++) {
                
                var opt = document.createElement('option');
                opt.innerHTML = arrayport[pp];
                opt.value = arrayport[pp];
                selpor.appendChild(opt);
  
            }


        }
    }
  

}




