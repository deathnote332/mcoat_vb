Public Class Dashboard

    Dim closeit As Integer

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        If MsgBox("Are you sure you want to logout?" & vbCrLf & fullname, vbQuestion + vbYesNo, "System") = vbYes Then
            Me.Hide()
            Login.Show()

        End If

    End Sub

    Private Sub Form37_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Enabled = True

        closeit = 0

        Label1.Text = fullname

        If userType = 1 Then

            Button14.Enabled = True
            Button15.Enabled = True

        Else

            Button14.Enabled = False
            Button15.Enabled = False

        End If


    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        Label7.Text = DateAndTime.Now

    End Sub


    Private Sub Button1_MouseHover(sender As Object, e As EventArgs) Handles btnSearch.MouseHover

        Label10.Text = "The user has the ability to search product by category, description, code and brand"

    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles btnSearch.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover

        Label10.Text = "The user has the ability to add, edit and delete products"

    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button3_MouseEnter(sender As Object, e As EventArgs) Handles btnProductout.MouseEnter

        Label10.Text = " The user can issued official receipt"

    End Sub

    Private Sub Button3_MouseLeave(sender As Object, e As EventArgs) Handles btnProductout.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button4_MouseHover(sender As Object, e As EventArgs) Handles btnProductin.MouseHover

        Label10.Text = " The user can update the stocks coming from the supplier"

    End Sub

    Private Sub Button4_MouseLeave(sender As Object, e As EventArgs) Handles btnProductin.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button5_MouseHover(sender As Object, e As EventArgs) Handles Button5.MouseHover

        Label10.Text = " The user can view the remaining stocks in the warehouse"

    End Sub

    Private Sub Button5_MouseLeave(sender As Object, e As EventArgs) Handles Button5.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button6_MouseHover(sender As Object, e As EventArgs) Handles Button6.MouseHover

        Label10.Text = " The user can modify and delete official receipt"

    End Sub

    Private Sub Button6_MouseLeave(sender As Object, e As EventArgs) Handles Button6.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button7_MouseHover(sender As Object, e As EventArgs) Handles Button7.MouseHover

        Label10.Text = " The user has the ability to track receipts"

    End Sub

    Private Sub Button7_MouseLeave(sender As Object, e As EventArgs) Handles Button7.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button8_MouseHover(sender As Object, e As EventArgs) Handles Button8.MouseHover

        Label10.Text = " The user has the ability to view completion of receipts"

    End Sub

    Private Sub Button8_MouseLeave(sender As Object, e As EventArgs) Handles Button8.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button9_MouseHover(sender As Object, e As EventArgs) Handles Button9.MouseHover

        Label10.Text = " The user has the ability to retrieve the deleted receipts"

    End Sub

    Private Sub Button9_MouseLeave(sender As Object, e As EventArgs) Handles Button9.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button10_MouseHover(sender As Object, e As EventArgs) Handles Button10.MouseHover

        Label10.Text = " The user has the ability to track the products delivered to the branches"

    End Sub

    Private Sub Button10_MouseLeave(sender As Object, e As EventArgs) Handles Button10.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button11_MouseHover(sender As Object, e As EventArgs) Handles Button11.MouseHover

        Label10.Text = " The user has the ability to track the products delivered to the warehouse"

    End Sub

    Private Sub Button11_MouseLeave(sender As Object, e As EventArgs) Handles Button11.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button12_MouseHover(sender As Object, e As EventArgs) Handles Button12.MouseHover

        Label10.Text = " The user has the ability to issued a purchase order to the supplier"

    End Sub

    Private Sub Button12_MouseLeave(sender As Object, e As EventArgs) Handles Button12.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button13_MouseHover(sender As Object, e As EventArgs) Handles Button13.MouseHover

        Label10.Text = " The user has the ability to create a report according to the total stocks of the branches"

    End Sub

    Private Sub Button13_MouseLeave(sender As Object, e As EventArgs) Handles Button13.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button15_MouseHover(sender As Object, e As EventArgs) Handles Button15.MouseHover

        Label10.Text = " The administrator has the ability to view the things happened in the system"

    End Sub

    Private Sub Button15_MouseLeave(sender As Object, e As EventArgs) Handles Button15.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button14_MouseHover(sender As Object, e As EventArgs) Handles Button14.MouseHover

        Label10.Text = " The administrator has the ability to add and remove accounts"

    End Sub

    Private Sub Button14_MouseLeave(sender As Object, e As EventArgs) Handles Button14.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button18_MouseHover(sender As Object, e As EventArgs) Handles Button18.MouseHover

        Label10.Text = " The people created the system"

    End Sub

    Private Sub Button18_MouseLeave(sender As Object, e As EventArgs) Handles Button18.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button17_MouseHover(sender As Object, e As EventArgs) Handles Button17.MouseHover

        Label10.Text = "The user can save and print the product being delivered"

    End Sub

    Private Sub Button17_MouseLeave(sender As Object, e As EventArgs) Handles Button17.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button20_MouseHover(sender As Object, e As EventArgs) Handles Button20.MouseHover

        Label10.Text = "The user can create reports such as monthly order, additional order and etc"

    End Sub

    Private Sub Button20_MouseLeave(sender As Object, e As EventArgs) Handles Button20.MouseLeave

        Label10.Text = ""

    End Sub

    Private Sub Button21_MouseHover(sender As Object, e As EventArgs) Handles Button21.MouseHover

        Label10.Text = "The user can view products according to Fast Moving, Slow Moving and Non Moving"

    End Sub

    Private Sub Button21_MouseLeave(sender As Object, e As EventArgs) Handles Button21.MouseLeave

        Label10.Text = ""

    End Sub



    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        SearchProduct.ShowDialog()
    End Sub

    Private Sub btnProductout_Click(sender As Object, e As EventArgs) Handles btnProductout.Click
        Productout.ShowDialog()
    End Sub
End Class