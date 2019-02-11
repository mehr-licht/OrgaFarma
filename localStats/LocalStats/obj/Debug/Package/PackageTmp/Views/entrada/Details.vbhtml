@ModelType LocalStats.input
@Code
    ViewData("Title") = "Details"
End Code

<h2>Details</h2>

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
</div>
<p>
    @Html.ActionLink("Edit", "Edit", New With { .id = Model.ID }) |
    @Html.ActionLink("Back to List", "Index")
</p>
