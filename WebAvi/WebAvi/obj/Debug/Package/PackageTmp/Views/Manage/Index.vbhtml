@ModelType IndexViewModel
@Code
    ViewBag.Title = "Gestão de conta"
End Code

<h2>@ViewBag.Title.</h2>

<p class="text-success">@ViewBag.StatusMessage</p>
<div>
    <h4>Alterar as definições da sua conta</h4>
    <hr />
    <dl class="dl-horizontal">
        <dt>Palavra passe</dt>
        <dd>
            [
            @If Model.HasPassword Then
                @Html.ActionLink("Altere a sua palavra passe", "ChangePassword")
            Else
                @Html.ActionLink("Criar", "SetPassword")
            End If
            ]
        </dd>
      @*  <dt>External Logins:</dt>
    <dd>
        @Model.Logins.Count [
        @Html.ActionLink("Gestão", "ManageLogins") ]
    </dd>

        Phone Numbers can used as a second factor of verification in a two-factor authentication system.

         See <a href="http://go.microsoft.com/fwlink/?LinkId=403804">this article</a>
            for details on setting up this ASP.NET application to support two-factor authentication using SMS.

         Uncomment the following block after you have set up two-factor authentication
    *@
        @* 
            <dt>Phone Number:</dt>
            <dd>
                @(If(Model.PhoneNumber, "None")) [
                @If (Model.PhoneNumber <> Nothing) Then
                    @Html.ActionLink("Change", "AddPhoneNumber")
                    @: &nbsp;|&nbsp;
                    @Html.ActionLink("Remove", "RemovePhoneNumber")
                Else
                    @Html.ActionLink("Add", "AddPhoneNumber")
                End If
                ]
            </dd>
       
        <dt>Two-Factor Authentication:</dt>
        <dd>
            <p>
                There are no two-factor authentication providers configured. See <a href="http://go.microsoft.com/fwlink/?LinkId=403804">this article</a>
                for details on setting up this ASP.NET application to support two-factor authentication.
            </p>
           
            @If Model.TwoFactor Then
                @Using Html.BeginForm("DisableTwoFactorAuthentication", "Manage", FormMethod.Post, New With { .class = "form-horizontal", .role = "form" })
                  @Html.AntiForgeryToken()
                  @<text>
                  Enabled
                  <input type="submit" value="Disable" class="btn btn-link" />
                  </text>
                End Using
            Else
                @Using Html.BeginForm("EnableTwoFactorAuthentication", "Manage", FormMethod.Post, New With { .class = "form-horizontal", .role = "form" })
                  @Html.AntiForgeryToken()
                  @<text>
                  Disabled
                  <input type="submit" value="Enable" class="btn btn-link" />
                  </text>
                End Using
            End If

    </dd>  *@
    </dl>
</div>
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