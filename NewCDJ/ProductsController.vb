Module ProductsController

    Const ID As Integer = 0
    Const BRAND As Integer = 1
    Const CATEGORY As Integer = 2
    Const DESCRIPTION As Integer = 3
    Const CODE As Integer = 4
    Const UNIT As Integer = 5
    Const QUANTITY As Integer = 6
    Const UNIT_PRICE As Integer = 7

    Private rs1 As New ADODB.Recordset
   
    Public Function search(searchString As String, dg As DataGridView, cb As ComboBox, Optional getAll As Boolean = False)

        If (getAll = True) Then
            rs.Open("select * from tblproducts order by description", connection, 2, 2)

            dg.Rows.Clear()
            dg.Visible = True

            Do Until rs.EOF

                dg.Rows.Add(rs("ID").Value, rs("brand").Value, rs("category").Value, rs("code").Value, rs("description").Value, rs("quantity").Value, rs("unit").Value, rs("unit_price").Value, rs("quantity").Value * rs("unit_price").Value)

                rs.MoveNext()

            Loop
            rs.Close()
        Else

        If cb.Text = "Brand" Then

            rs.Open("select * from tblproducts where brand like '%" & searchString & "%' order by description", connection, 2, 2)

            dg.Rows.Clear()
            dg.Visible = True

            Do Until rs.EOF

                dg.Rows.Add(rs("ID").Value, rs("brand").Value, rs("category").Value, rs("code").Value, rs("description").Value, rs("quantity").Value, rs("unit").Value, rs("unit_price").Value, rs("quantity").Value * rs("unit_price").Value)

                rs.MoveNext()

            Loop
                rs.Close()
        ElseIf cb.Text = "Category" Then

            rs.Open("select * from tblproducts where category like '%" & searchString & "%' order by brand,description", connection, 2, 2)

            dg.Rows.Clear()
            dg.Visible = True

            Do Until rs.EOF

                dg.Rows.Add(rs("ID").Value, rs("brand").Value, rs("category").Value, rs("code").Value, rs("description").Value, rs("quantity").Value, rs("unit").Value, rs("unit_price").Value, rs("quantity").Value * rs("unit_price").Value)

                rs.MoveNext()

            Loop
                rs.Close()
        ElseIf cb.Text = "Description" Then

            rs.Open("select * from tblproducts where description like '%" & searchString & "%' order by brand,category", connection, 2, 2)

            dg.Rows.Clear()
            dg.Visible = True

            Do Until rs.EOF

                dg.Rows.Add(rs("ID").Value, rs("brand").Value, rs("category").Value, rs("code").Value, rs("description").Value, rs("quantity").Value, rs("unit").Value, rs("unit_price").Value, rs("quantity").Value * rs("unit_price").Value)

                rs.MoveNext()

            Loop
                rs.Close()
        ElseIf cb.Text = "Code" Then

            rs.Open("select * from tblproducts where code like '%" & searchString & "%' order by brand", connection, 2, 2)

            dg.Rows.Clear()
            dg.Visible = True

            Do Until rs.EOF

                dg.Rows.Add(rs("ID").Value, rs("brand").Value, rs("category").Value, rs("code").Value, rs("description").Value, rs("quantity").Value, rs("unit").Value, rs("unit_price").Value, rs("quantity").Value * rs("unit_price").Value)

                rs.MoveNext()

            Loop
                rs.Close()
        End If

        End If

        For Each rw As DataGridViewRow In dg.Rows

            If rw.Cells("column6").Value <= 3 And rw.Cells("column6").Value >= 1 Then

                rw.DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48)

            ElseIf rw.Cells("column6").Value = 0 Then

                rw.DefaultCellStyle.BackColor = Color.FromArgb(28, 28, 28)

            ElseIf rw.Cells("column6").Value >= 4 Then

                rw.DefaultCellStyle.BackColor = Color.FromArgb(104, 33, 122)

            End If

        Next

        If dg.Rows.Count = 0 Then

            dg.Visible = False

        End If
        Return Nothing
    End Function


    Public Function insertUpdateTemp(product_id As Integer, qty As Integer, old_qty As Integer, Optional updateQty As Integer = 0, Optional insert_update As Integer = 1)

        If (insert_update = 1) Then
            rs.Open("select * from temp_product_out", connection, 2, 2)
            rs.AddNew()
            rs("product_id").Value = product_id
            rs("qty").Value = qty
            rs.Update()
            rs.Close()
        Else
            rs.Open("select * from temp_product_out where product_id='" & product_id & "'", connection, 2, 2)
            rs("qty").Value = updateQty
            rs.Update()
            rs.Close()
        End If
        Dim newqty = old_qty - qty
        updateProduct(product_id, newqty)

        Return Nothing
    End Function

    Public Function deleteTemp(product_id As Integer, qty As Integer)
        Dim oldqty As Integer = getProductDetails(product_id)(QUANTITY)
        Dim newqty = oldqty + qty
        updateProduct(product_id, newqty)
        rs.Open("select * from temp_product_out where product_id='" & product_id & "'", connection, 2, 2)
        rs.Delete()
        rs.Close()
        Return Nothing
    End Function

    Private Function getProductDetails(product_id As Integer) As Object()

        rs.Open("select * from tblproducts where id='" & product_id & "'", connection, 2, 2)
        Dim newObject As Object() = New Object() {rs("id").Value, rs("brand").Value, rs("category").Value, rs("description").Value, rs("code").Value, rs("unit").Value, rs("quantity").Value, rs("unit_price").Value}
        rs.Close()
        Return newObject
    End Function

    Private Function updateProduct(product_id As Integer, newqty As Integer)
        rs.Open("select * from tblproducts where id='" & product_id & "'", connection, 2, 2)
        rs("quantity").Value = newqty
        rs.Update()
        rs.Close()
        Return Nothing
    End Function

    Private Function tableTemp(Optional limit As Integer = 28, Optional idProductID As Integer = 1)
        Dim arrayID As New List(Of String)
        Dim arrayProductID As New List(Of String)
        rs.Open("select * from temp_product_out limit " & limit, connection, 2, 2)
        Do Until rs.EOF
            arrayID.Add(rs("id").Value)
            arrayProductID.Add(rs("product_id").Value)
            rs.MoveNext()
        Loop
        rs.Close()


        If (idProductID = 1) Then
            Return arrayID.ToArray
        Else
            Return arrayProductID.ToArray
        End If

    End Function

    

    Public Function printProductout(rpt As CrystalReport1)

        Dim receipt_no As String = "MC-" & Now.Year & "-" & Format(getLastinsertedReciept, "0000#")
        insertIntoProductOut(receipt_no, getTempTotal)

        Dim str = Strings.Join(tableTemp(2, 2), ",")
        Form1.CrystalReportViewer1.ReportSource = Nothing
        rs.Open("select * from tblproducts where id in (" & str & ")", connection, 2, 2)
        rpt.SetDataSource(rs)
        rpt.PrintOptions.PrinterName = DefaultPrinterName()
        rpt.PrintToPrinter(nCopies:=1, collated:=0, endPageN:=0, startPageN:=0)
        rs.Close()

        deleteMultipleTemp(receipt_no)


        Return Nothing


    End Function

    Public Function deleteMultipleTemp(rec_no As String)

        'insert into product_out items
        Dim str1 = Strings.Join(tableTemp(2, 2), ",")
        Dim total As Double
        rs.Open("select temp_product_out.qty as temp_qty,tblproducts.* from temp_product_out join tblproducts on temp_product_out.product_id = tblproducts.id  where tblproducts.id in (" & str1 & ")", connection, 2, 2)
        Do Until rs.EOF
            insertIntoProductoutItem(rec_no, rs("id").Value, rs("temp_qty").Value)
            rs.MoveNext()
        Loop
        rs.Close()

        Dim str = Strings.Join(tableTemp(2, 1), ",")
        rs.Open("delete from temp_product_out where id in (" & str & ")", connection, 2, 2)
        Return Nothing
    End Function

    Public Function getTempTotal()
        Dim str = Strings.Join(tableTemp(2, 2), ",")
        Dim total As Double
        rs.Open("select sum(temp_product_out.qty * tblproducts.unit_price) as total from tblproducts join temp_product_out on temp_product_out.product_id = tblproducts.id  where tblproducts.id in (" & str & ")", connection, 2, 2)
        total = rs("total").Value
        rs.Close()

        Return total
    End Function

    Private Function insertIntoProductoutItem(rec_no As String, product_id As Integer, qty As Integer)

        rs1.Open("select * from product_out_items", connection, 2, 2)
        rs1.AddNew()
        rs1("receipt_no").Value = rec_no
        rs1("product_id").Value = product_id
        rs1("quantity").Value = qty
        rs1.Update()
        rs1.Close()

        Return Nothing
    End Function

   
    Private Function insertIntoProductOut(rec_no As String, total As Double)

        rs.Open("select * from product_out", connection, 2, 2)
        rs.AddNew()
        rs("receipt_no").Value = rec_no
        rs("total").Value = total
        rs.Update()
        rs.Close()

        Return Nothing
    End Function

    Private Function getLastinsertedReciept()

        Dim lastID As Integer
        rs.Open("select * from product_out order by id desc", connection, 2, 2)
        If rs.EOF Then
            lastID = 0
        Else
            lastID = rs("id").Value
        End If

        rs.Close()

        Return Val(lastID + 1)

    End Function


    Public Function getTempProductout(listView As ListView)

        rs.Open("select temp_product_out.qty as temp_qty,tblproducts.* from temp_product_out join tblproducts on temp_product_out.product_id = tblproducts.id", connection, 2, 2)
        Do Until rs.EOF
            Dim lvt As New ListViewItem

            lvt.SubItems.Add(Val(listView.Items.Count) + 1)

            lvt.SubItems.Add(rs("category").Value)
            lvt.SubItems.Add(rs("description").Value)
            lvt.SubItems.Add(rs("brand").Value)
            lvt.SubItems.Add(rs("unit").Value)
            lvt.SubItems.Add(rs("temp_qty").Value)
            lvt.SubItems.Add(rs("unit_price").Value)
            lvt.SubItems.Add(Val(rs("temp_qty").Value) * Val(rs("unit_price").Value))
            lvt.SubItems.Add(rs("id").Value)

            listView.Items.Add(lvt)

            rs.MoveNext()
        Loop
        rs.Close()


        
        Return Nothing

    End Function


End Module
