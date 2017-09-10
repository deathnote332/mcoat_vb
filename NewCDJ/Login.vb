Public Class Login

    Dim down As Integer

    Private Sub Form9_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        down = 4

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        If txtUsername.Text = "" Or txtPassword.Text = "" Then

            MsgBox("Please input a username or password first!", vbExclamation, "Error")

            txtUsername.Focus()

        Else
            If (authenticate(txtUsername.Text, txtPassword.Text)) Then
                Me.Hide()
                Dashboard.Show()
            Else
                down = down - 1

                If down = 0 Then

                    MsgBox("Three (3) attempts tried!" & vbCrLf & "System will shutdown", vbCritical, "Error")

                    End

                End If

                MsgBox("Access denied!" & vbCrLf & "Remaining attempts: " & down, vbExclamation, "Error")

                txtUsername.Text = ""
                txtUsername.Focus()

                txtPassword.Text = ""
            End If

        End If

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown

        If e.KeyCode = Keys.Enter Then

            btnLogin.PerformClick()

        End If

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress

        If Char.IsLetterOrDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Then

            e.Handled = False

        Else

            e.Handled = True

        End If

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUsername.KeyPress

        If Char.IsLetterOrDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Then

            e.Handled = False

        Else

            e.Handled = True

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Timer1.Enabled = False
        Timer2.Enabled = False

        Application.Exit()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If down = 0 Then

            Timer1.Enabled = False

            MsgBox("Three (3) attempts tried!" & vbCrLf & "System will shutdown", vbCritical, "Error")

            End

        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Label3.BackColor = Color.Transparent
        Label3.Text = Now.ToLongDateString & "   " & Now.ToLongTimeString

    End Sub


End Class