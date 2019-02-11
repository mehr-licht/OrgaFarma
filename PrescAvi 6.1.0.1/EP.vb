Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System
Imports System.Windows.Forms
Imports System.Security.Permissions
'Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb


Public Class EP
    Inherits Form


    Private Sub EP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo MOSTRARERRO
        Me.KeyPreview = True

        Exit Sub
MOSTRARERRO:
        MsgBox("Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub
End Class