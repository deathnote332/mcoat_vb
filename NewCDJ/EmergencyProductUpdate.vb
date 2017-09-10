
Public Class EmergencyProductUpdate

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If txtBrand.Text = "" Or txtCategory.Text = "" Or txtCode.Text = "" Or txtDescription.Text = "" Or txtQuantity.Text = "" Or txtUnit.Text = "" Then

            MsgBox("Please fill-up all the information first!", vbExclamation, "Error")

            txtBrand.Focus()

        Else

    

        End If

        Me.Close()

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

        Me.Close()

    End Sub

End Class