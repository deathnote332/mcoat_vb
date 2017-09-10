Imports CrystalDecisions.Shared.TableLogOnInfo
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing

Public Class Productout

    Dim basis As String
    Dim crSections As Sections
    Dim crSection As Section
    Dim crReportObjects As ReportObjects
    Dim crReportObject As ReportObject
    Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
    Dim myLogin As CrystalDecisions.Shared.TableLogOnInfo
    Dim crSubreportObject As SubreportObject
    Dim crSubreportDocument As ReportDocument
    Dim crDatabase As Database
    Dim crTables As Tables
    Dim aTable As Table
    Dim crTableLogOnInfo As TableLogOnInfo
    Dim repDoc As New CrystalReport1 'Result is the report doc .rpt
    Dim a As String
    Dim sto As String

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.F1 And TextBox1.Enabled = True Then

            TextBox1.Text = ""
            TextBox1.Focus()

        ElseIf e.KeyCode = Keys.F1 And TextBox1.Enabled = False Then

            ComboBox2.Text = ""
            ComboBox2.Focus()

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        getTempProductout(ListView1)
        countPrintPages()

        cbBranches.Text = ""
        ComboBox2.Text = ""
        cbBranches.Focus()

        TextBox1.Text = ""
        txtID.Text = ""
        txtBrand.Text = ""
        txtCategory.Text = ""
        txtDescription.Text = ""
        txtCode.Text = ""
        txtUnitPrice.Text = ""
        txtUnit.Text = ""
        txtCurrentQty.Text = ""
        txtInputRelease.Text = ""
        TextBox12.Text = ""
        'TextBox11.Text = ""

        Label17.Text = "No. of receipt()"

        TextBox1.Enabled = False

        DataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48)
        DataGridView1.DefaultCellStyle.ForeColor = Color.White

        getBranches(cbBranches)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged


        If TextBox1.Text = "" Or ComboBox2.Text = "" Then

            DataGridView1.Rows.Clear()
            DataGridView1.Visible = False

        Else

            search(TextBox1.Text, DataGridView1, ComboBox2)

        End If



    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs)

        If e.KeyCode = Keys.Enter Then

            Call Button2.PerformClick()

        End If

    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then

            e.Handled = True

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If txtInputRelease.Text = "" Or Val(txtInputRelease.Text) = 0 Then

            MsgBox("Please input a quantity first!", vbExclamation, "Error")

            txtInputRelease.Text = ""
            txtInputRelease.Focus()

        ElseIf Val(txtInputRelease.Text) > Val(txtCurrentQty.Text) Then

            MsgBox("Not enough stock!", vbExclamation, "Error")

            txtInputRelease.Text = ""
            txtInputRelease.Focus()

        ElseIf txtDescription.Text = "" Then

            MsgBox("Please search first!", vbExclamation, "Error")

            TextBox1.Text = ""
            TextBox1.Focus()

        Else

            Dim newquantity As Integer

            newquantity = Val(txtCurrentQty.Text) - Val(txtInputRelease.Text)

            Dim lvt As ListViewItem
            Dim lcount As Integer
            Dim flag As Integer

            lvt = New ListViewItem

            lcount = 0
            flag = 0

            TextBox11.Text = ""
            TextBox12.Text = ""

            Do Until lcount = ListView1.Items.Count

                If ListView1.Items(lcount).SubItems(9).Text = txtID.Text Then

                    Dim newq As Double

                    newq = 0
                    flag = flag + 1

                    newq = Val(ListView1.Items(lcount).SubItems(6).Text) + Val(txtInputRelease.Text)

                    ListView1.Items(lcount).SubItems(6).Text = Val(newq)
                    ListView1.Items(lcount).SubItems(8).Text = Val(txtUnitPrice.Text) * newq

                    TextBox12.Text = Val(txtUnitPrice.Text) * newq

                    insertUpdateTemp(Val(txtID.Text), Val(txtInputRelease.Text), Val(txtCurrentQty.Text), Val(newq), 2)


                End If

                TextBox11.Text = Val(ListView1.Items(lcount).SubItems(8).Text) + Val(TextBox11.Text)

                lcount = lcount + 1

            Loop

            If flag = 0 Then

                lvt.SubItems.Add(Val(ListView1.Items.Count) + 1)

                lvt.SubItems.Add(txtCategory.Text)
                lvt.SubItems.Add(txtDescription.Text)
                lvt.SubItems.Add(txtBrand.Text)
                lvt.SubItems.Add(txtUnit.Text)
                lvt.SubItems.Add(Val(txtInputRelease.Text))
                lvt.SubItems.Add(txtUnitPrice.Text)
                lvt.SubItems.Add(Val(txtInputRelease.Text) * Val(txtUnitPrice.Text))
                lvt.SubItems.Add(txtID.Text)

                ListView1.Items.Add(lvt)

                lvt.Selected = True
                lvt.EnsureVisible()


                TextBox11.Text = Val(TextBox11.Text) + (Val(txtInputRelease.Text) * Val(txtUnitPrice.Text))
                TextBox12.Text = Val(txtInputRelease.Text) * Val(txtUnitPrice.Text)

                insertUpdateTemp(Val(txtID.Text), Val(txtInputRelease.Text), txtCurrentQty.Text)


            End If

            countPrintPages()
            End If

            txtID.Text = ""
            txtBrand.Text = ""
            txtCategory.Text = ""
            txtDescription.Text = ""
            txtCode.Text = ""
            txtCurrentQty.Text = ""
            txtUnit.Text = ""
            txtUnitPrice.Text = ""
            txtInputRelease.Text = ""

            Button1.Enabled = True
            txtInputRelease.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = True

            cbBranches.Enabled = True

            TextBox1.Text = ""
            TextBox1.Focus()

        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        txtID.Text = DataGridView1.CurrentRow.Cells(0).Value
        txtBrand.Text = DataGridView1.CurrentRow.Cells(1).Value
        txtCategory.Text = DataGridView1.CurrentRow.Cells(2).Value
        txtDescription.Text = DataGridView1.CurrentRow.Cells(3).Value
        txtCode.Text = DataGridView1.CurrentRow.Cells(4).Value
        txtCurrentQty.Text = DataGridView1.CurrentRow.Cells(5).Value
        txtUnit.Text = DataGridView1.CurrentRow.Cells(6).Value
        txtUnitPrice.Text = DataGridView1.CurrentRow.Cells(7).Value

        DataGridView1.Visible = False

        TextBox1.Text = ""
        txtInputRelease.Text = ""
        txtInputRelease.Focus()

        txtInputRelease.Enabled = True
        Button2.Enabled = True

    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown

        If e.KeyCode = Keys.Enter Then

            txtID.Text = DataGridView1.CurrentRow.Cells(0).Value
            txtBrand.Text = DataGridView1.CurrentRow.Cells(1).Value
            txtCategory.Text = DataGridView1.CurrentRow.Cells(2).Value
            txtDescription.Text = DataGridView1.CurrentRow.Cells(3).Value
            txtCode.Text = DataGridView1.CurrentRow.Cells(4).Value
            txtCurrentQty.Text = DataGridView1.CurrentRow.Cells(5).Value
            txtUnit.Text = DataGridView1.CurrentRow.Cells(6).Value
            txtUnitPrice.Text = DataGridView1.CurrentRow.Cells(7).Value

            DataGridView1.Visible = False


            TextBox1.Text = ""
            txtInputRelease.Text = ""
            txtInputRelease.Focus()

            txtInputRelease.Enabled = True
            Button2.Enabled = True

        End If

        If e.KeyCode = Keys.F8 Then

            'Form39.TextBox9.Text = DataGridView1.CurrentRow.Cells(1).Value
            'Form39.TextBox10.Text = DataGridView1.CurrentRow.Cells(2).Value
            'Form39.TextBox11.Text = DataGridView1.CurrentRow.Cells(4).Value
            'Form39.TextBox12.Text = DataGridView1.CurrentRow.Cells(3).Value
            'Form39.TextBox13.Text = DataGridView1.CurrentRow.Cells(5).Value
            'Form39.TextBox14.Text = DataGridView1.CurrentRow.Cells(6).Value
            'Form39.TextBox15.Text = DataGridView1.CurrentRow.Cells(7).Value

            'Form39.TextBox17.Text = DataGridView1.CurrentRow.Cells(1).Value
            'Form39.TextBox18.Text = DataGridView1.CurrentRow.Cells(2).Value
            'Form39.TextBox19.Text = DataGridView1.CurrentRow.Cells(4).Value
            'Form39.TextBox20.Text = DataGridView1.CurrentRow.Cells(3).Value
            'Form39.TextBox21.Text = DataGridView1.CurrentRow.Cells(5).Value
            'Form39.TextBox22.Text = DataGridView1.CurrentRow.Cells(6).Value
            'Form39.TextBox23.Text = DataGridView1.CurrentRow.Cells(7).Value

            'DataGridView1.Visible = False

            'Form39.ShowDialog()

        End If

    End Sub

    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click

      
        If MsgBox("Are you sure you want to delete this product?" & vbCrLf & ListView1.FocusedItem.SubItems(2).Text & " - " & ListView1.FocusedItem.SubItems(3).Text & " - " & ListView1.FocusedItem.SubItems(4).Text & " - " & ListView1.FocusedItem.SubItems(5).Text, vbQuestion + vbYesNo, "System") = vbYes Then

            MsgBox("Product deleted from the query", vbInformation, "System")

            Dim product_id As Integer = Val(ListView1.FocusedItem.SubItems(9).Text)
            Dim product_qty As Integer = Val(Val(ListView1.FocusedItem.SubItems(6).Text))

            deleteTemp(product_id, product_qty)

            ListView1.FocusedItem.Remove()
            Call recount()
            countPrintPages()

            txtID.Text = ""
            txtBrand.Text = ""
            txtCategory.Text = ""
            txtDescription.Text = ""
            txtCode.Text = ""
            txtCurrentQty.Text = ""
            txtUnit.Text = ""
            txtUnitPrice.Text = ""
            txtInputRelease.Text = ""
            TextBox1.Text = ""

            If txtID.Text = "" Then

                TextBox1.Text = ""
                TextBox1.Focus()

            Else

                txtInputRelease.Text = ""
                txtInputRelease.Focus()

            End If

        End If

        Exit Sub


    End Sub

    Private Sub ListView1_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles ListView1.ColumnWidthChanging

        Dim DisableColumns As Integer() = {1, 2, 3, 4, 5, 6, 7, 8}

        For Each DCol As Integer In DisableColumns

            If e.ColumnIndex = DCol Then

                e.Cancel = True
                e.NewWidth = ListView1.Columns(DCol).Width

            End If

        Next DCol

    End Sub

    Private Sub ListView1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView1.KeyDown

        If e.KeyCode = Keys.F4 Then

            On Error GoTo errhandling

            If MsgBox("Are you sure you want to delete this product?" & vbCrLf & ListView1.FocusedItem.SubItems(2).Text & " - " & ListView1.FocusedItem.SubItems(3).Text & " - " & ListView1.FocusedItem.SubItems(4).Text & " - " & ListView1.FocusedItem.SubItems(5).Text, vbQuestion + vbYesNo, "System") = vbYes Then

                MsgBox("Product deleted from the query", vbInformation, "System")


                Call recount()
                countPrintPages()
                
                txtID.Text = ""
                txtBrand.Text = ""
                txtCategory.Text = ""
                txtDescription.Text = ""
                txtCode.Text = ""
                txtCurrentQty.Text = ""
                txtUnit.Text = ""
                txtUnitPrice.Text = ""
                txtInputRelease.Text = ""
                TextBox1.Text = ""

                If txtID.Text = "" Then

                    TextBox1.Text = ""
                    TextBox1.Focus()

                Else

                    txtInputRelease.Text = ""
                    txtInputRelease.Focus()

                End If

            End If

        End If

        Exit Sub

errhandling:

        MsgBox("There is nothing to delete", vbExclamation, "Error")

        If txtID.Text = "" Then

            TextBox1.Text = ""
            TextBox1.Focus()

        Else

            txtInputRelease.Text = ""
            txtInputRelease.Focus()

        End If

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click

        If cbBranches.Text = "" Then

            MsgBox("Please choose a destination to be sent first!" & vbCrLf & vbCrLf, vbExclamation, "Error")

            cbBranches.Focus()

        Else

            If MsgBox("Are you sure you want to print?", vbQuestion + vbYesNo, "System") = vbYes Then

                Dim rpt As New CrystalReport1

                Dim os As Integer
                Dim nom As Integer

                os = 0
                nom = 0

                Do Until os = ListView1.Items.Count

                    If (os Mod 2) = 0 Then

                        nom = nom + 1

                        Label17.Text = "No. of receipt(" & nom & ")"

                    End If

                    os = os + 1

                Loop

               

                If nom = 0 Then

                    Label17.Text = "No. of receipt(1)"

                End If

                Do Until nom = 0
                    '   printProductout(rpt)

                    nom = nom - 1
                Loop



            End If

        End If

    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbBranches.KeyPress

        e.Handled = True

    End Sub

    Private Sub ComboBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox2.KeyPress

        e.Handled = True

    End Sub

    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged

        basis = ComboBox2.Text

        TextBox1.Enabled = True

        TextBox1.Text = ""
        TextBox1.Focus()

        DataGridView1.Visible = False

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)

        ' Form4.ShowDialog()

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbBranches.SelectedValueChanged

        sto = ""
        sto = cbBranches.Text

    End Sub

    Private Function recount()

        Dim renum As Integer

        renum = 1

        For c = 0 To ListView1.Items.Count - 1

            ListView1.Items(c).SubItems(1).Text = renum

            renum = renum + 1

        Next

        Return Nothing

    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        If MsgBox("Are you sure you want to void the process?", vbYesNo + vbQuestion, "System") = vbYes Then

           

        End If

    End Sub


    

    Private Sub Button6_Click(sender As Object, e As EventArgs)

        '  Form7.ShowDialog()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)

        Me.Close()

    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click

        ComboBox2.Text = "Brand"
        TextBox1.Text = ""

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        ComboBox2.Text = "Category"
        TextBox1.Text = ""

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        ComboBox2.Text = "Description"
        TextBox1.Text = ""

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        ComboBox2.Text = "Code"
        TextBox1.Text = ""

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)

        '  Form42.ShowDialog()

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

        Me.Close()

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click

        ' Form4.ShowDialog()

    End Sub

    Private Sub TextBox10_KeyDown1(sender As Object, e As KeyEventArgs) Handles txtInputRelease.KeyDown

        If e.KeyCode = Keys.Enter Then

            Button2.PerformClick()
            Timer1.Enabled = True

        End If

    End Sub

    Private Sub TextBox10_KeyPress1(sender As Object, e As KeyPressEventArgs) Handles txtInputRelease.KeyPress

        If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then

            e.Handled = True

        End If
    End Sub

    Private Function countPrintPages()
        Dim os As Integer
        Dim nom As Integer

        os = 0
        nom = 0

        Do Until os = ListView1.Items.Count

            If (os Mod 28) = 0 Then

                nom = nom + 1

                Label17.Text = "No. of receipt(" & nom & ")"

            End If

            os = os + 1

        Loop

        If nom = 0 Then

            Label17.Text = "No. of receipt(1)"

        End If
        Return nom
    End Function
End Class
