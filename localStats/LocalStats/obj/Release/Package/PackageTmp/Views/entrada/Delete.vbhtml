@ModelType LocalStats.input
@Code
    ViewData("Title") = "Delete"
End Code

<h2>Delete</h2>

<h3>Are you sure you want to delete this?</h3>
<div>
    <h4>input</h4>
    <hr />
    <dl class="dl-horizontal">
        <dt>
            @Html.DisplayNameFor(Function(model) model.utente)
        </dt>

        <dd>
            @Html.DisplayFor(Function(model) model.utente)
        </dd>


        <dt>
            @Html.DisplayNameFor(Function(model) model.local)
        </dt>

        <dd>
            @Html.DisplayFor(Function(model) model.local)
        </dd>

        
        <dt>
            @Html.DisplayNameFor(Function(model) model.qty)
        </dt>

        <dd>
            @Html.DisplayFor(Function(model) model.qty)
        </dd>

    </dl>
    @Using (Html.BeginForm())
        @Html.AntiForgeryToken()

        @<div class="form-actions no-color">
            <input type="submit" value="Delete" class="btn btn-default" /> |
            @Html.ActionLink("Back to List", "Index")
        </div>
    End Using
</div>
