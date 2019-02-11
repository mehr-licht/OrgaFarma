@ModelType IEnumerable(Of LocalStats.input)
@Code
ViewData("Title") = "Index"
End Code

<h2>Index</h2>

<p>
    @Html.ActionLink("Create New", "Create")
</p>
<table class="table">
    <tr>
       
        <th>
            @Html.DisplayNameFor(Function(model) model.utente)
        </th>
        <th>
            @Html.DisplayNameFor(Function(model) model.local)
        </th>

        <th>
            @Html.DisplayNameFor(Function(model) model.qty)
        </th>
        <th></th>
    </tr>

@For Each item In Model
    @<tr>
         <td>
             @Html.DisplayFor(Function(modelItem) item.utente)
         </td>
        <td>
            @Html.DisplayFor(Function(modelItem) item.local)
        </td>
        
        <td>
            @Html.DisplayFor(Function(modelItem) item.qty)
        </td>
        <td>
            @Html.ActionLink("Edit", "Edit", New With {.id = item.ID }) |
            @Html.ActionLink("Details", "Details", New With {.id = item.ID }) |
            @Html.ActionLink("Delete", "Delete", New With {.id = item.ID })
        </td>
    </tr>
Next

</table>
