Imports System
Imports System.Data.Entity.Migrations
Imports Microsoft.VisualBasic

Namespace Migrations
    Public Partial Class InitialCreate
        Inherits DbMigration
    
        Public Overrides Sub Up()
            CreateTable(
                "dbo.inputs",
                Function(c) New With
                    {
                        .ID = c.Int(nullable := False, identity := True),
                        .local = c.Int(nullable := False),
                        .utente = c.Int(nullable := False),
                        .qty = c.Int(nullable := False)
                    }) _
                .PrimaryKey(Function(t) t.ID)
            
        End Sub
        
        Public Overrides Sub Down()
            DropTable("dbo.inputs")
        End Sub
    End Class
End Namespace
