

//function movetoNext(current, nextFieldID) {

//    if (current.value.length >= current.maxLength) {
//        document.getElementById(nextFieldID).focus();
//        if (nextFieldID == "a2") {

//            document.getElementById("btnSubmit").style.visibility = "visible";
//        }
//    }
//}




//    function naofocar() {
//        var focado2 = document.getElementById(document.activeElement.id);
//        var fim2 = focado2.id.substring(1);
//        var inicio2 = focado2.id.substring(0, 1);
//        var anterior2 = inicio2 + (fim2 - 1);
//        if (focado2.value.length == 0) {
//            if (fim2 != 1){
//                document.getElementById(anterior2).focus();
//            }
//        }
//    }

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



$("input[type=text]").keydown(function (e) {
    if (e.which == 13 || e.which == 9) {

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
        if (seguinte == "p5") {
            seguinte = "a1";
        }

        if (focado.value.length == 0) { //enter no vazio
            if (fim != 1) { //se não é p1 nem a1
                if (document.getElementById(anterior).value.length == 0) { //se anterior vazio
                    document.getElementById(anterior).focus(); //foca no anterior

                } else { //se anterior não vazio
                    if (seguinte == "a5") { //se é a4
                        document.getElementById("btnSubmit").click(); //foca no seguinte

                    } else { //se nºao é a4
                        switch (inicio) {
                            case "p":
                                document.getElementById("a1").focus();//enter no p vazio  <> p1 vai focar no a1
                                break;
                            case "a":
                                document.getElementById("btnSubmit").click();//enter no a vazio  <> a1 vai clickar botão
                        }


                    }
                }
            } else { //é p1 ou a1
                document.getElementById(focado.id).focus(); //foca no próprio

            }
        } else {
            if (focado.value.length < 7 || (isNaN(focado.value))) { //focado incompleto ou com simbolos/letras
                document.getElementById(focado.id).innerHtml = ""; //limpa o focado vazio

                if (anterior.value.length == 0) {
                    document.getElementById(anterior).focus(); //foca no anterior
                    document.getElementById(anterior).trigger(e);//manda enter para voltar a testar anterior
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
                        document.getElementById("btnSubmit").style.visibility = "visible";//se enter no a1 comparar fica visivel
                    }
                }
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

    $("#btnSubmit").click(function () {
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
                        if (labeltotal = "l") {
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
                            if (blue[1][1] != "a") {
                                $("#labellinha1").css("fontSize", "20px");
                                $("#labellinha1").css("fontWeight", "bold");
                                $("#labellinha1").css("color", (blue[1][5]));
                                $("#labellinha1").css("text-shadow", "#000 1px 1px 1px");
                                $("#labellinha1").css("-webkit - font - smoothing", "antialiased");
                                $("#labeltotal1").html(blue[1][0]);
                                var qual1p = (blue[1][0]).substring(0, 1);
                                $("#p" + qual1p).css("background-color", (blue[1][5]));
                                $("#a1").css("background-color", (blue[1][5]));
                                tooltipalinea("labellinha1", (blue[1][1]));
                                tooltipalinea("labeltotal1", (blue[1][1]));
                                //if (blue[1][1] == "w") {                              cnpemfillajax não funciona por isso...
                                //    $("#showRight").click();
                                //    $("#cnpemquery").text = $("#a1").val;
                                //    querycnpem();
                                //}
                            }
                        }

                        if (blue[2][0] != null) {
                            if (blue[2][1] != "a") {
                                $("#labellinha2").css("fontSize", "20px");
                                $("#labellinha2").css("fontWeight", "bold");
                                $("#labellinha2").css("color", "red");
                                $("#labellinha2").css("text-shadow", "#000 1px 1px 1px");
                                $("#labellinha2").css("-webkit - font - smoothing", "antialiased");
                                $("#labeltotal2").html(blue[2][0]);
                                var qual2p = (blue[2][0]).substring(0, 1);
                                $("#p" + qual2p).css("background-color", (blue[2][5]));
                                $("#a2").css("background-color", (blue[2][5]));
                                tooltipalinea("labellinha2", (blue[2][1]));
                                tooltipalinea("labeltotal2", (blue[2][1]));
                                if (blue[2][1] == "w") {
                                }
                            }
                        }

                        if (blue[3][0] != null) {
                            if (blue[3][1] != "a") {
                                $("#labellinha3").css("fontSize", "20px");
                                $("#labellinha3").css("fontWeight", "bold");
                                $("#labellinha3").css("color", "red");
                                $("#labellinha3").css("text-shadow", "#000 1px 1px 1px");
                                $("#labellinha3").css("-webkit - font - smoothing", "antialiased");
                                $("#labeltotal3").html(blue[3][0]);
                                var qual3p = (blue[3][0]).substring(0, 1);
                                $("#p" + qual3p).css("background-color", (blue[3][5]));
                                $("#a3").css("background-color", (blue[3][5]));
                                tooltipalinea("labellinha3", (blue[3][1]));
                                tooltipalinea("labeltotal3", (blue[3][1]));
                                if (blue[3][1] == "w") {
                                }
                            }
                        }

                        if (blue[4][0] != null) {
                            if (blue[4][1] != "a") {
                                $("#labellinha4").css("fontSize", "20px");
                                $("#labellinha4").css("fontWeight", "bold");
                                $("#labellinha4").css("color", "red");
                                $("#labellinha4").css("text-shadow", "#000 1px 1px 1px");
                                $("#labellinha4").css("-webkit - font - smoothing", "antialiased");
                                $("#labeltotal4").html(blue[4][0]);
                                var qual4p = (blue[4][0]).substring(0, 1);
                                $("#p" + qual4p).css("background-color", (blue[4][5]));
                                $("#a4").css("background-color", (blue[4][5]));
                                tooltipalinea("labellinha4", (blue[4][1]));
                                tooltipalinea("labeltotal4", (blue[4][1]));
                                if (blue[4][1] == "w") {
                                }
                            }
                        }
                        //if (blue[0][2] == true){
                        //    $('#labeltotalports').html("despacho")
                        //}
                        //if (blue[0][3] == true) {
                        //    $('#labeltotalexcep').html("excepção")
                        //}
                        //if (blue[0][4] == true) {
                        //    $('#labeltotaltop5').html("top 5")
                        //}
                        $('#labellinha1').html(blue[1][0]);
                        $('#labellinha2').html(blue[2][0]);
                        $('#labellinha3').html(blue[3][0]);
                        $('#labellinha4').html(blue[4][0]);

                        var corcor = blue[0][5];
                        mudarimagem(corcor);


                        if (blue[12][0] > 0) {
                            $('#a4').prop('title', 'cnpem: ' + blue[12][15] + '\r\n' + 'nome: ' + blue[12][2] + '\r\n' + 'dci: ' + blue[12][1] + '\r\n' + 'forma: ' + blue[12][3] + '\r\n' + 'dosagem: ' + blue[12][4] + '\r\n' + 'x ' + blue[12][5]);// + '\r\n' + 'comp= ' + blue[12][6] + '%' + '\r\n' + 'GH: ' + blue[12][7] + '\r\n' + 'pvp= €' + blue[12][8] + '\r\n' + 'pr= €' + blue[12][9] + '\r\n' + 'top5= €' + blue[12][12] + '\r\n' + 'lab= ' + blue[12][11] + '\r\n' + 'gen= ' + blue[12][10] + '\r\n' + 'por dci= ' + blue[12][16]);
                        }
                        if (blue[11][0] > 0) {
                            $('#a3').prop('title', 'cnpem: ' + blue[11][15] + '\r\n' + 'nome: ' + blue[11][2] + '\r\n' + 'dci: ' + blue[11][1] + '\r\n' + 'forma: ' + blue[11][3] + '\r\n' + 'dosagem: ' + blue[11][4] + '\r\n' + 'x ' + blue[11][5]);// + '\r\n' + 'comp= ' + blue[11][6] + '%' + '\r\n' + 'GH: ' + blue[11][7] + '\r\n' + 'pvp= €' + blue[11][8] + '\r\n' + 'pr= €' + blue[11][9] + '\r\n' + 'top5= €' + blue[11][12] + '\r\n' + 'lab= ' + blue[11][11] + '\r\n' + 'gen= ' + blue[11][10] + '\r\n' + 'por dci= ' + blue[11][16]);
                        }
                        if (blue[10][0] > 0) {
                            $('#a2').prop('title', 'cnpem: ' + blue[10][15] + '\r\n' + 'nome: ' + blue[10][2] + '\r\n' + 'dci: ' + blue[10][1] + '\r\n' + 'forma: ' + blue[10][3] + '\r\n' + 'dosagem: ' + blue[10][4] + '\r\n' + 'x ' + blue[10][5]);// + '\r\n' + 'comp= ' + blue[10][6] + '%' + '\r\n' + 'GH: ' + blue[10][7] + '\r\n' + 'pvp= €' + blue[10][8] + '\r\n' + 'pr= €' + blue[10][9] + '\r\n' + 'top5= €' + blue[10][12] + '\r\n' + 'lab= ' + blue[10][11] + '\r\n' + 'gen= ' + blue[10][10] + '\r\n' + 'por dci= ' + blue[10][16]);
                        }
                        if (blue[9][0] > 0) {
                            $('#a1').prop('title', 'cnpem: ' + blue[9][15] + '\r\n' + 'nome: ' + blue[9][2] + '\r\n' + 'dci: ' + blue[9][1] + '\r\n' + 'forma: ' + blue[9][3] + '\r\n' + 'dosagem: ' + blue[9][4] + '\r\n' + 'x ' + blue[9][5]);// + '\r\n' + 'comp= ' + blue[9][6] + '%' + '\r\n' + 'GH: ' + blue[9][7] + '\r\n' + 'pvp= €' + blue[9][8] + '\r\n' + 'pr= €' + blue[9][9] + '\r\n' + 'top5= €' + blue[9][12] + '\r\n' + 'lab= ' + blue[9][11] + '\r\n' + 'gen= ' + blue[9][10] + '\r\n' + 'por dci= ' + blue[9][16]);
                        }

                        if (blue[8][0] > 0) {
                            $('#p4').prop('title', 'cnpem: ' + blue[8][15] + '\r\n' + 'nome: ' + blue[8][2] + '\r\n' + 'dci: ' + blue[8][1] + '\r\n' + 'forma: ' + blue[8][3] + '\r\n' + 'dosagem: ' + blue[8][4] + '\r\n' + 'x ' + blue[8][5]);// + '\r\n' + 'comp= ' + blue[8][6] + '%' + '\r\n' + 'GH: ' + blue[8][7] + '\r\n' + 'pvp= €' + blue[8][8] + '\r\n' + 'pr= €' + blue[8][9] + '\r\n' + 'top5= €' + blue[8][12] + '\r\n' + 'lab= ' + blue[8][11] + '\r\n' + 'gen= ' + blue[8][10] + '\r\n' + 'por dci= ' + blue[8][16]);
                        }
                        if (blue[7][0] > 0) {
                            $('#p3').prop('title', 'cnpem: ' + blue[7][15] + '\r\n' + 'nome: ' + blue[7][2] + '\r\n' + 'dci: ' + blue[7][1] + '\r\n' + 'forma: ' + blue[7][3] + '\r\n' + 'dosagem: ' + blue[7][4] + '\r\n' + 'x ' + blue[7][5]);// + '\r\n' + 'comp= ' + blue[7][6] + '%' + '\r\n' + 'GH: ' + blue[7][7] + '\r\n' + 'pvp= €' + blue[7][8] + '\r\n' + 'pr= €' + blue[7][9] + '\r\n' + 'top5= €' + blue[7][12] + '\r\n' + 'lab= ' + blue[7][11] + '\r\n' + 'gen= ' + blue[7][10] + '\r\n' + 'por dci= ' + blue[7][16]);
                        }
                        if (blue[6][0] > 0) {
                            $('#p2').prop('title', 'cnpem: ' + blue[6][15] + '\r\n' + 'nome: ' + blue[6][2] + '\r\n' + 'dci: ' + blue[6][1] + '\r\n' + 'forma: ' + blue[6][3] + '\r\n' + 'dosagem: ' + blue[6][4] + '\r\n' + 'x ' + blue[6][5]);// + '\r\n' + 'comp= ' + blue[6][6] + '%' + '\r\n' + 'GH: ' + blue[6][7] + '\r\n' + 'pvp= €' + blue[6][8] + '\r\n' + 'pr= €' + blue[6][9] + '\r\n' + 'top5= €' + blue[6][12] + '\r\n' + 'lab= ' + blue[6][11] + '\r\n' + 'gen= ' + blue[6][10] + '\r\n' + 'por dci= ' + blue[6][16]);
                        }
                        if (blue[5][0] > 0) {
                            $('#p1').prop('title', 'cnpem: ' + blue[5][15] + '\r\n' + 'nome: ' + blue[5][2] + '\r\n' + 'dci: ' + blue[5][1] + '\r\n' + 'forma: ' + blue[5][3] + '\r\n' + 'dosagem: ' + blue[5][4] + '\r\n' + 'x' + blue[5][5]);// + '\r\n' + 'comp= ' + blue[5][6] + '%' + '\r\n' + 'GH: ' + blue[5][7] + '\r\n' + 'pvp= €' + blue[5][8] + '\r\n' + 'pr= €' + blue[5][9] + '\r\n' + 'top5= €' + blue[5][12] + '\r\n' + 'lab= ' + blue[5][11] + '\r\n' + 'gen= ' + blue[5][10] + '\r\n' + 'por dci= ' + blue[5][16]);
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




    });

});


function mudarimagem(colour) {

    var image = document.getElementById('semaforo');
    image.src = "../img/" + colour + ".gif"

    document.getElementById("btnlimpar").style.visibility = "visible"
    document.getElementById("btnlimpar").focus()

};







function toggle() {
    document.getElementById("btnSubmit").style.visibility = "hidden"

}


$(window).keypress(function (e) {

    if (e.which == 32) {

        document.getElementById("btnlimpar").click();
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

    document.getElementById("labeltotal0").setAttribute('title', "");
    document.getElementById("labeltotal0").innerHTML = "";
    document.getElementById("labeltotal1").innerHTML = "";
    document.getElementById("labeltotal2").innerHTML = "";
    document.getElementById("labeltotal3").innerHTML = "";
    document.getElementById("labeltotal4").innerHTML = "";
    //document.getElementById("labeltotal5").innerHTML = "";
    //document.getElementById("labeltotal6").innerHTML = "";
    document.getElementById("p1").style.backgroundColor = "lightblue";
    document.getElementById("p2").style.backgroundColor = "lightblue";
    document.getElementById("p3").style.backgroundColor = "lightblue";
    document.getElementById("p4").style.backgroundColor = "lightblue";
    document.getElementById("a1").style.backgroundColor = "lightblue";
    document.getElementById("a2").style.backgroundColor = "lightblue";
    document.getElementById("a3").style.backgroundColor = "lightblue";
    document.getElementById("a4").style.backgroundColor = "lightblue";

};









$(document).ready(function () {
    $(cnpemquery).keypress(function (e) {
        if (e.which == 13) {

            var param = $('#cnpemquery').val();
         
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: "../prescavi/cnpemfill",
                data: "{pedido:'" + param + "'}",
                //dataType: "text",
                success: function (dev) {

                 
                    //var dev = $.parseJSON(data);

                    debugger
                    if (dev != null) {
                        alert("retornou");
                        alert("jquery" + dev[1][0]);
                        if (dev.length > 0) {

                            if (dev[0].length > 0) {

                                $('#labelcnpem1').html(dev[0][0]);
                                $('#labelcnpem1').prop('title', 'nome: ' + dev[0][2] + '\r\n' + 'dci: ' + dev[0][1] + '\r\n' + 'forma: ' + dev[0][3] + '\r\n' + 'dosagem: ' + dev[0][4] + '\r\n' + 'x ' + dev[0][5] + '\r\n' + 'lab= ' + dev[0][11]);


                            } else {
                                //linhacnpem1 = não foram encontrados códigos que satisfaçam este cnpem
                            }
                        }
                    } //fim blue != null
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



            });//fecha ajax
        }//fecha if
    });
});
