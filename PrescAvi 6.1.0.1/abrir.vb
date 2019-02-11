Public Class abrir

   
    'Dim form1loadado As Boolean
    'Dim soalgarismosA As Boolean
    'Dim focoA As String
    'Dim farmaciaA As Integer


    'Private Sub GetOpenFormTitles()
    'Dim formTitles As New Collection
    '
    '   Try
    '      For Each f As Form In My.Application.OpenForms
    '         If Not f.InvokeRequired Then
    ' Can access the form directly.
    '            formTitles.Add(f.Text)
    '       End If
    '  Next
    '     Catch ex As Exception
    '        formTitles.Add("Error: " & ex.Message)
    '   End Try
    '  For i = 0 To 3
    '     If formTitles(i) Is Form1 Then
    '        form1loadado = True
    '   End If
    '      Next
    ' End Sub

    'Private Sub abrir_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Me.KeyPreview = True
    '    GetOpenFormTitles()
    '   If form1loadado = True Then
    'MsgBox(form1loadado)
    'End If
    'MsgBox(form1loadado)
    'End Sub




    ' Private Sub frmDesigner_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
    '    On Error GoTo MOSTRARERRO
    '   focoA = Me.ActiveControl.Name()
    '
    'Dim soalgarismos As Boolean
    '   If Not Char.IsNumber(e.KeyChar) Then
    '      soalgarismosA = False
    ' Else
    '    soalgarismosA = True
    'End If

    '    Select Case soalgarismosA
    '       Case False
    '          Beep()
    '
    '    End Select


    '     Select Case focoA
    '        Case "abrirtextbox"
     '           Me.abrirtextbox.BackColor = Color.White
    '          Select Case soalgarismosA
    '             Case False
    '                Beep()
    '               Me.abrirtextbox.BackColor = Color.Red
    '          Case True
    '             If Asc(e.KeyChar) = Keys.Enter Then
    '                Select Case focoA
    '                   Case "abrirtextbox"
    '                      If abrirtextbox.Text = "" Then
    '                         Beep()
    '                           Me.abrirtextbox.BackColor = Color.Red
    'Else
    '  Select Case abrirtextbox.Text
    '       Case Is = 8168
    '            farmaciaA = 8168
    '             fecharabrireabrirform1()
    '          Case Is = 19542
    '               farmaciaA = 19542
    '                fecharabrireabrirform1()
    '             Case Is = 4243
    '                  farmaciaA = 4243
    '                   fecharabrireabrirform1()
    '                Case Is = 1813
    '                     farmaciaA = 1813
    '                      fecharabrireabrirform1()
    '                   Case Is = 3441
    '    farmaciaA = 3441
    '     fecharabrireabrirform1()
    '  Case Is = 3948
    '       farmaciaA = 3948
    '        fecharabrireabrirform1()
    '     Case Is = 14400
    '          farmaciaA = 14400
    '           fecharabrireabrirform1()
    '        Case Else
    '             Me.abrirtextbox.BackColor = Color.Red
    '              Beep()
    '               Me.abrirtextbox.Focus()
    '                abrirtextbox.SelectionStart = 0
    '                 abrirtextbox.SelectionLength = Len(abrirtextbox.Text)
    '                  End Select
    '                                 End If
    '                           End Select
    '                  End If
    '
    '           End Select
    '
    '
    '        End Select
    '
    '
    '
    '
    '
    '   Exit Sub
    'MOSTRARERRO:
    '       MsgBox("Sub frmDesigner_KeyPress: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '      Resume Next
    ' End Sub


    ' Sub fecharabrireabrirform1()
    '     On Error GoTo MOSTRARERRO
    ' Dim obj As New Form1
    ' ' obj.PassedText = abrirtextbox.Text
    '    If form1loadado = False Then
    '        obj.Show()
    '    End If
    '    form1loadado = True
    '    Me.Close()
    'MOSTRARERRO:
    '       MsgBox("SUB fecharabrireabrirform1: Erro número #" & Str$(Err.Number) & " na linha " & Str$(Erl) & " - " & Err.Description & " - gerado por " & Err.Source)
    '       Resume Next
    '  End Sub


    'Private Sub abrirbutao_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles abrirbutao.Click
    '    Me.abrirtextbox.BackColor = Color.White
    '
    '
    '   If abrirtextbox.Text = "" Then
    '      Beep()
    '     Me.abrirtextbox.BackColor = Color.Red
    'Else
    '   Select Case abrirtextbox.Text
    '      Case Is = 19542
    '         farmaciaA = 19542
    '        fecharabrireabrirform1()
    '   Case Is = 8168
    '      farmaciaA = 8168
    '     fecharabrireabrirform1()
    'Case Is = 4243
    '    farmaciaA = 4243
    '    fecharabrireabrirform1()
    ' Case Is = 1813
    '     farmaciaA = 1813
    '     fecharabrireabrirform1()
    ' Case Is = 3441
    '    farmaciaA = 3441
    '     fecharabrireabrirform1()
    ' Case Is = 3948
    '     farmaciaA = 3948
    '     fecharabrireabrirform1()
    ' Case Is = 14400
    '     farmaciaA = 14400
    '     fecharabrireabrirform1()
    ' Case Else
    '     Me.abrirtextbox.BackColor = Color.Red
    '     Beep()
    '     Me.abrirtextbox.Focus()
    '     abrirtextbox.SelectionStart = 0
    '     abrirtextbox.SelectionLength = Len(abrirtextbox.Text)
    '     End Select
    '     End If
    '

    ' End Sub
End Class