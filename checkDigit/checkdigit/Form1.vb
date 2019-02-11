Imports System.Text

Public Class Form1


    Dim s As String
    Dim n As Double
    Dim r As Double
    Dim caractp1 As String
    Dim valorponto As String
    Dim valor As String
    Dim numero As String
    Dim check As String
    Dim aviados As String
    Dim number As String
    Dim foco As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        limpar()
    End Sub

    Private Sub frmDesigner_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
1:      foco = Me.ActiveControl.Name()
    End Sub





    Private Sub data_KeysEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        If e.KeyCode = Keys.Enter Then

            Select Case foco
                Case Is = "numerotxtbx"
41:                 If numerotxtbx.Text <> "" Then
                        Me.aviadotxtbx.Focus()
                    End If
42:             Case Is = "aviadotxtbx"
                    If aviadotxtbx.Text <> "" Then
                        Me.verifbtn.Focus()
                        Me.verifbtn.Select()
                    Else
                        Me.aviadotxtbx.Focus()
                    End If
                Case Is = "verifbtn"
                    clickar()
                Case Is = "form1"
                    limpar()
                Case Else
                    numerotxtbx.Focus()
            End Select
        End If
    End Sub

    Sub clickar()
        On Error GoTo MOSTRARERRO

        If numerotxtbx.Text = "" Or aviadotxtbx.Text = "" Then
            limpar()
            Exit Sub
        End If
2:      number = numerotxtbx.Text
3:      aviados = aviadotxtbx.Text
4:      check = number.Substring(Len(number) - 1, 1).ToUpper
5:      'ordem = number.Substring(0, Len(number) - 1)
6:      numero = getNumeric(number.Substring(0, Len(number) - 2))
      
7:      s = aviados & numero
8:
9:      n = Val(s)
10:     While n > 26
11:         s = 0
12:         While n > 0
13:             r = n Mod 10
14:             s = s + r
15:             n = n \ 10
16:         End While
17:         n = s
18:     End While
19:     valor = numero & aviados & letrar(s)
20:     valorponto = numero & "." & aviados & letrar(s)
       
21:     If valor = UCase(numerotxtbx.Text) Or valorponto = UCase(numerotxtbx.Text) Then
22:         If MessageBox.Show("OK", "OK", MessageBoxButtons.OKCancel) Then

23:             limpar()
                numerotxtbx.Focus()
24:         End If
25:     Else
26:         If MessageBox.Show("Não Bate Certo", "ERRO!", MessageBoxButtons.OKCancel) Then

27:             limpar()
                numerotxtbx.Focus()
28:         End If
29:     End If
30:     Exit Sub
MOSTRARERRO:
        MsgBox("Sub Clickar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Sub limpar()
1:      On Error GoTo MOSTRARERRO
        caractp1 = ""
2:      valor = ""
3:      numero = ""
4:      check = ""
5:      aviados = ""
6:      number = ""
7:      numerotxtbx.Text = ""
8:      aviadotxtbx.Text = ""
        Me.numerotxtbx.Focus()

9:      Exit Sub
MOSTRARERRO:
        MsgBox("Sub limpar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Private Sub verifbtn_Click(sender As Object, e As EventArgs) Handles verifbtn.Click
1:      On Error GoTo MOSTRARERRO
2:      clickar()
        numerotxtbx.Focus()
30:     Exit Sub
MOSTRARERRO:
        MsgBox("Sub verifbtn_Click: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Sub



    Public Function letrar(ByVal valor As Integer)
1:      On Error GoTo MOSTRARERRO
2:      letrar = Chr(64 + valor)
        Exit Function
MOSTRARERRO:
        MsgBox("function letrar: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
        Resume Next
    End Function


    Public Function getNumeric(value As String) As String
        Dim output As StringBuilder = New StringBuilder
        For i = 0 To value.Length - 1
            If IsNumeric(value(i)) Then
                output.Append(value(i))
            End If
        Next
        Return output.ToString()
    End Function


    Sub focoaviado() Handles aviadotxtbx.GotFocus
        If aviadotxtbx.Text = "" And numerotxtbx.Text = "" Then
            numerotxtbx.Focus()
        End If
    End Sub


    Private Sub aviadotxtbx_TextChanged(sender As Object, e As EventArgs) Handles aviadotxtbx.TextChanged
        If Len(aviadotxtbx.Text) = 1 Then
            caractp1 = aviadotxtbx.Text
            If caractp1 Like "#" Then
                clickar()
                numerotxtbx.Focus()
            End If
        End If
        If aviadotxtbx.Text = "" And numerotxtbx.Text = "" Then
            numerotxtbx.Focus()
        End If
    End Sub
End Class
