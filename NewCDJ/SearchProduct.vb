Public Class SearchProduct

   
    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress

        e.Handled = True

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text = "" Or ComboBox1.Text = "" Then

            DataGridView1.Rows.Clear()
            DataGridView1.Visible = False

        Else
            search(TextBox1.Text, DataGridView1, ComboBox1)

        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        If ComboBox1.Text = "All" Then

            TextBox1.Text = ""
            TextBox1.Visible = False

            search(TextBox1.Text, DataGridView1, ComboBox1, True)

        ElseIf ComboBox1.Text = "Brand" Or ComboBox1.Text = "Category" Or ComboBox1.Text = "Code" Or ComboBox1.Text = "Description" Then
            DataGridView1.Rows.Clear()

            DataGridView1.Visible = False
            TextBox1.Visible = True
            TextBox1.Enabled = True

            TextBox1.Text = ""
            TextBox1.Focus()

        End If

    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Text = ""
        ComboBox1.Focus()

        TextBox1.Text = ""


    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        Me.Close()

    End Sub

End Class