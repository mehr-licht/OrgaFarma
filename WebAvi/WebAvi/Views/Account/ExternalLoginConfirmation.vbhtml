@ModelType ExternalLoginConfirmationViewModel
@Code
    ViewBag.Title = "Registar"
End Code

<h2>@ViewBag.Title.</h2>
<h3>Associar a sua @ViewBag.LoginProvider conta.</h3>

@Using Html.BeginForm("ExternalLoginConfirmation", "Account", New With { .ReturnUrl = ViewBag.ReturnUrl }, FormMethod.Post, New With {.class = "form-horizontal", .role = "form"})
    @Html.AntiForgeryToken()

    @<text>
    <h4></h4>
    <hr />
    @Html.ValidationSummary(True, "", New With {.class = "text-danger"})
    <p class="text-info">
      Autenticou-se com sucesso como <strong>@ViewBag.LoginProvider</strong>.
       Por favor introduza um nome de usuário e click no botão de registo para completar o registo.
       
    <div class="form-group">
        @Html.LabelFor(Function(m) m.Email, New With {.class = "col-md-2 control-label"})
        <div class="col-md-10">
            @Html.TextBoxFor(Function(m) m.Email, New With {.class = "form-control"})
            @Html.ValidationMessageFor(Function(m) m.Email, "", New With {.class = "text-danger"})
        </div>
    </div>
    <div class="form-group">
        <div class="col-md-offset-2 col-md-10">
            <input type="submit" class="btn btn-default" value="Registar" />
        </div>
    </div>
    </text>
End Using

@Section Scripts
    @Scripts.Render("~/bundles/jqueryval")
End Section
        <style>
          BODY {background-color: cadetblue;}
          </style>