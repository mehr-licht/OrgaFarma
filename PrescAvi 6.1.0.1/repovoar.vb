Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb

Public Class repovoar

    Sub repovoar_load()

    End Sub


    


    Private Sub ImportarPvpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportarPvpToolStripMenuItem.Click
        Dim objDataAdapter As New SqlDataAdapter("Select * From pvp", "pvp.mdb")
        Dim dsResult As New DataSet("Result")

        If Not IsNothing(objDataAdapter) Then
            ' Fill data into dataset
            objDataAdapter.Fill(dsResult)

            objDataAdapter.Dispose()
        End If
    End Sub


End Class