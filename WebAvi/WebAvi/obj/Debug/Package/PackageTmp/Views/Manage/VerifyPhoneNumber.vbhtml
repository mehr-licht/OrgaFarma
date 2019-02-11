@ModelType VerifyPhoneNumberViewModel
@Code
    ViewBag.Title = "Verificar o número de telefone"
End Code

<h2>@ViewBag.Title.</h2>

@Using Html.BeginForm("VerifyPhoneNumber", "Manage", FormMethod.Post, New With { .class = "form-horizontal", .role = "form" })
    @Html.AntiForgeryToken()
    @Html.Hidden("phoneNumber", Model.PhoneNumber)
    @<text>
    <h4>Introduza o code de verificação</h4>
    <h5>@ViewBag.Status</h5>
    <hr />
    @Html.ValidationSummary("", New With { .class = "text-danger" })
    <div class="form-group">
        @Html.LabelFor(Function(m) m.Code, New With { .class = "col-md-2 control-label" })
        <div class="col-md-10">
            @Html.TextBoxFor(Function(m) m.Code, New With { .class = "form-control" })
        </div>
    </div>
    <div class="form-group">
        <div class="col-md-offset-2 col-md-10">
            <input type="submit" class="btn btn-default" value="OK" />
        </div>
    </div>
    </text>
End Using

@Section Scripts
    @Scripts.Render("~/bundles/jqueryval")
End Section
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