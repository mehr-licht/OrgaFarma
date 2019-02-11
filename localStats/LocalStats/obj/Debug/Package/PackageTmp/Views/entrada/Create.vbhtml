@ModelType LocalStats.input
@Code
    ViewData("Title") = "Create"
  
    
End Code

    <h2>Create</h2>

    @Using (Html.BeginForm())
        @Html.AntiForgeryToken()

        @<div class="form-horizontal">
            <h4>input</h4>
            <hr />
            @Html.ValidationSummary(True, "", New With {.class = "text-danger"})


             <div class="form-group">
                 @Html.LabelFor(Function(model) model.local, htmlAttributes:=New With {.class = "control-label col-md-2"})
                 <div class="col-md-10">
                     @Html.TextBoxFor(Function(model) model.local, New With {.htmlAttributes = New With {.class = "form-control"}})
                     @Html.ValidationMessageFor(Function(model) model.local, "", New With {.class = "text-danger"})
                 </div>
             </div>

            <div class="form-group">
                @Html.LabelFor(Function(model) model.utente, htmlAttributes:=New With {.class = "control-label col-md-2"})
                <div class="col-md-10">
                    @Html.TextBoxFor(Function(model) model.utente, New With {.htmlAttributes = New With {.class = "form-control"}})
                    
                    @Html.ValidationMessageFor(Function(model) model.utente, "", New With {.class = "text-danger"})
                </div>
            </div>
    

            <div class="form-group">
                @Html.LabelFor(Function(model) model.qty, htmlAttributes:=New With {.class = "control-label col-md-2"})
                <div class="col-md-10">
                    @Html.TextBoxFor(Function(model) model.qty, New With {.htmlAttributes = New With {.class = "form-control"}})
                    @Html.ValidationMessageFor(Function(model) model.qty, "", New With {.class = "text-danger"})
                </div>
            </div>

            <div class="form-group">
                <div class="col-md-offset-2 col-md-10">
                    <input type="submit" value="Create" class="btn btn-default" />
                </div>
            </div>
             <div class="form-group">
                 <div class="col-md-offset-2 col-md-10">
                     <input type="button" value="Clear" class="btn btn-default" onclick="Create" />
                 </div>
             </div>
        </div>
    End Using

    <div>
        @Html.ActionLink("Back to List", "Index")
    </div>
   
    @Section Scripts
        @Scripts.Render("~/bundles/jqueryval")


    End Section
