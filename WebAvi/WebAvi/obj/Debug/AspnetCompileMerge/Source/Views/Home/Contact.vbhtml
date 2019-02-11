@Code
    ViewData("Title") = "Contactos"
End Code

<h2>@ViewData("Title").</h2>
<style>
    .forma{
        width:200px;
    }
</style>

<form action="mailto:webavi@orgafarma.com" method="GET">
    <div class="forma">
    <input name="subject" type="text" value="assunto"/>
    <textarea name="body"></textarea>
    <input type="submit" value="enviar" />
        </div>
</form>
<address>
 <br />
    <br />
    <abbr title="telemóvel">Telf:</abbr>
   91.590.7334
</address>

<address>
    <strong>Apoio:</strong>   <a href="mailto:apoio@orgafarma.pt" class="coisas">apoio@orgafarma.pt</a><br />
    <strong>informações:</strong> <a href="mailto:webavi@orgafarma.com" class="coisas">webavi@orgafarma.pt</a>
</address>
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