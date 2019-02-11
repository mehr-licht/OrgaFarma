$(document).ready(function () {
    var menuLeft = document.getElementById('cbp-spmenu-s1'),
        				menuBottom = document.getElementById('cbp-spmenu-s4'),
                            menuRight = document.getElementById('cbp-spmenu-s2'),
				showLeft = document.getElementById('showLeft'),
                showRight = document.getElementById('showRight'),
				showBottom = document.getElementById('showBottom'),

                                body = document.body;

    showLeft.onclick = function () {
        datar()
        classie.toggle(this, 'active');
        classie.toggle(menuLeft, 'cbp-spmenu-open');
        disableOther('showLeft');
    };
    showRight.onclick = function () {
        classie.toggle(this, 'active');
        classie.toggle(menuRight, 'cbp-spmenu-open');
        disableOther('showRight');
    };
    showBottom.onclick = function () {
        classie.toggle(this, 'active');
        classie.toggle(menuBottom, 'cbp-spmenu-open');
        disableOther('showBottom');
    };

    function highlightar(dia) {

        if (dia > 9) {
            document.getElementById("dia" + dia).style.fontWeight = "bold";
            document.getElementById("dia" + dia).style.color = "green";
            document.getElementById("dia0" + dia).style.textShadow = "#000 1px 1px 1px";
            document.getElementById("dia0" + dia).style.setProperty("-webkit - font - smoothing", "antialiased");
            
        } else {

            document.getElementById("dia0" + dia).style.fontWeight = "bold";
            document.getElementById("dia0" + dia).style.color = "green";
            document.getElementById("dia0" + dia).style.textShadow = "#000 1px 1px 1px";
            document.getElementById("dia0" + dia).style.setProperty("-webkit - font - smoothing", "antialiased");
                    }
    };

    function datar() {

        var hoje = new Date();
        highlightar(hoje.getUTCDate());
        var mes = hoje.getMonth() + 1;
        var year = hoje.getFullYear();
        var trinta = "";
        var renovavel = mes - 6;
        var meses = "/" + renovavel;
        var mesantes = mes - 1;
        switch (mes) {
            case 2:
                trinta = "-1";
                document.getElementById("dia30").style.visibility = "hidden";
                if (year != 2016 && year != 2020) {
                    document.getElementById("dia29").style.visibility = "hidden";
                }
                break;
            case 3:
                trinta = "+2";
                if (year == 2016 || year == 2020) {
                    trinta = "+1";
                }
                break;
            case 4:
            case 6:
            case 9:
            case 11:

                document.getElementById("dia31").style.visibility = "hidden";
                trinta = "-1";
                break;
            default:
                document.getElementById('dia01').textContent = "1     1/4";

                //getElementbyId("dia01").innerhtml = getElementbyId("dia01").getattribute("id").substring(3) + "  " + getElementbyId("dia01").getattribute("id").substring(3) + meses;
        }

        switch (mes) {
            case 1:
            case 3:
            case 5:
            case 7:
            case 8:
            case 10:
            case 12:
                mesantes = mes;

        }

        if (mes == 3) {
            document.getElementById("dia01").text = document.getElementById("dia01").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia01").id.substring(3)) + trinta) + "/" + mes + "......" + document.getElementById("dia01").id.substring(3) + meses;
        } else {
            document.getElementById("dia01").text = document.getElementById("dia01").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia01").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia01").id.substring(3) + meses;
        }


        if (mes == 3) {
            if (year == 2016 || year == 2020) {
                document.getElementById("dia02").text = document.getElementById("dia02").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia02").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia02").id.substring(3) + meses;
            }
        } else {
            document.getElementById("dia02").text = document.getElementById("dia02").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia02").id.substring(3)) + trinta) + "/" + mes + "......" + document.getElementById("dia02").id.substring(3) + meses;
        }

        document.getElementById("dia03").text = document.getElementById("dia03").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia03").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia03").id.substring(3) + meses;
        document.getElementById("dia04").text = document.getElementById("dia04").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia04").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia04").id.substring(3) + meses;
        document.getElementById("dia05").text = document.getElementById("dia05").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia05").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia05").id.substring(3) + meses;
        document.getElementById("dia06").text = document.getElementById("dia06").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia06").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia06").id.substring(3) + meses;
        document.getElementById("dia07").text = document.getElementById("dia07").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia07").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia07").id.substring(3) + meses;
        document.getElementById("dia08").text = document.getElementById("dia08").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia08").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia08").id.substring(3) + meses;
        document.getElementById("dia09").text = document.getElementById("dia09").id.substring(3) + "......0" + parseFloat(parseFloat(document.getElementById("dia09").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia09").id.substring(3) + meses;
        document.getElementById("dia10").text = document.getElementById("dia10").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia10").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia10").id.substring(3) + meses;
        document.getElementById("dia11").text = document.getElementById("dia11").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia11").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia11").id.substring(3) + meses;
        document.getElementById("dia12").text = document.getElementById("dia12").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia12").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia12").id.substring(3) + meses;
        document.getElementById("dia13").text = document.getElementById("dia13").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia13").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia13").id.substring(3) + meses;
        document.getElementById("dia14").text = document.getElementById("dia14").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia14").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia14").id.substring(3) + meses;
        document.getElementById("dia15").text = document.getElementById("dia15").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia15").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia15").id.substring(3) + meses;
        document.getElementById("dia16").text = document.getElementById("dia16").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia16").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia16").id.substring(3) + meses;
        document.getElementById("dia17").text = document.getElementById("dia17").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia17").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia17").id.substring(3) + meses;
        document.getElementById("dia18").text = document.getElementById("dia18").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia18").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia18").id.substring(3) + meses;
        document.getElementById("dia19").text = document.getElementById("dia19").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia19").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia19").id.substring(3) + meses;
        document.getElementById("dia20").text = document.getElementById("dia20").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia20").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia20").id.substring(3) + meses;
        document.getElementById("dia21").text = document.getElementById("dia21").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia21").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia21").id.substring(3) + meses;
        document.getElementById("dia22").text = document.getElementById("dia22").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia22").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia22").id.substring(3) + meses;
        document.getElementById("dia23").text = document.getElementById("dia23").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia23").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia23").id.substring(3) + meses;
        document.getElementById("dia24").text = document.getElementById("dia24").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia24").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia24").id.substring(3) + meses;
        document.getElementById("dia25").text = document.getElementById("dia25").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia25").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia25").id.substring(3) + meses;
        document.getElementById("dia26").text = document.getElementById("dia26").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia26").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia26").id.substring(3) + meses;
        document.getElementById("dia27").text = document.getElementById("dia27").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia27").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia27").id.substring(3) + meses;
        document.getElementById("dia28").text = document.getElementById("dia28").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia28").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia28").id.substring(3) + meses;
        document.getElementById("dia29").text = document.getElementById("dia29").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia29").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia29").id.substring(3) + meses;
        document.getElementById("dia30").text = document.getElementById("dia30").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia30").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia30").id.substring(3) + meses;
        document.getElementById("dia31").text = document.getElementById("dia31").id.substring(3) + "......" + parseFloat(parseFloat(document.getElementById("dia31").id.substring(3)) + trinta) + "/" + mesantes + "......" + document.getElementById("dia31").id.substring(3) + meses;
    }


})