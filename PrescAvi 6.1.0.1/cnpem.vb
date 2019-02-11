Imports System
Imports System.IO
Imports System.Security
'Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class cnpem

    Private Sub cnpem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        On Error GoTo MOSTRARERRO

        Me.KeyPreview = True
        cnpem_dci.Focus()
        Exit Sub
MOSTRARERRO:
        MsgBox("Sub cnpem_Load: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub


    Sub cnpem_dci_textchanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cnpem_dci.TextChanged

    End Sub


    Sub cnpem_dose_textchanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cnpem_dose.TextChanged

    End Sub

    Sub cnpem_forma_textchanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cnpem_forma.TextChanged

    End Sub

    Sub cnpem_qty_textchanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cnpem_qty.TextChanged

    End Sub

    Sub cnpem_but_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cnpem_but.Click

    End Sub

    Sub cnpem_but_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cnpem_but.Enter

    End Sub


    Public Sub sacarexcelmensal()
        Dim ficheiroxls As String
        ficheiroxls = "cnpem.xls"
        Dim connectionxls As OleDbConnection
        Dim connectionxlsstring As String
        Dim sqlxls As String
        connectionxlsstring = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & My.Computer.FileSystem.SpecialDirectories.Desktop & "/" & ficheiroxls & "; Extended Properties=""Excel 8.0;HDR=YES"";" 'hdr=yes quer dizer que tem cabeçalho
        sqlxls = "SELECT * FROM [Sheet1$]"
        connectionxls = New OleDbConnection(connectionxlsstring)
        Dim adapterxls As New OleDbDataAdapter(sqlxls, connectionxls)

        'adapterxls.Fill(medicamentosxls)
        MsgBox(ficheiroxls)
    End Sub




End Class