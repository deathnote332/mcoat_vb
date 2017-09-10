Imports System.Security.Cryptography
Imports System.Drawing.Printing

Module Globals

    'connection

    Public rs As New ADODB.Recordset
    Public rs1 As New ADODB.Recordset
    Public _hostname As String = "localhost"
    Public _username As String = "root"
    Public _password As String = ""
    Public _dbname As String = "mcoat_db"

    Public connection As String = "Driver=MySQL ODBC 3.51 Driver;Server=" & _hostname & ";Database=" & _dbname & ";Uid=" & _username & ";Pwd=" & _password & ";"

    Public connection1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\IS.mdb;Persist Security Info=True;Jet OLEDB:Database Password=cdjultimatepassword"



    'User details
    Public userType As String
    Public fullname As String



    'encryption
    Public trides As New TripleDESCryptoServiceProvider
    Public hasher As New MD5CryptoServiceProvider
    Public enc As Byte()


    Public Function getBranches(cb As ComboBox, Optional order_by As Integer = 1)
        Dim str As String
        If (order_by = 1) Then
            str = " order by name asc"
        Else
            str = " order by name desc"
        End If
        cb.Items.Clear()

        rs.Open("select * from branches" & str, connection, 2, 2)
        Do Until rs.EOF
            cb.Items.Add(New DictionaryEntry(rs("name").Value, rs("id").Value))
            rs.MoveNext()
        Loop
        rs.Close()

        cb.DisplayMember = "Key"
        cb.ValueMember = "Value"

    End Function

    Public Function DefaultPrinterName()

        Dim ps As New PrinterSettings()
        Return ps.PrinterName

    End Function


    Public Function encrypt(ByVal input As String)
        enc = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
        trides.Key = hasher.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes("mcoatpaintcommercial"))
        trides.Mode = CipherMode.ECB
        Return Convert.ToBase64String(trides.CreateEncryptor.TransformFinalBlock(enc, 0, enc.Length))
    End Function
    Public Function decrypt(ByVal input As String)
        enc = Convert.FromBase64String(input)
        trides.Key = hasher.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes("mcoatpaintcommercial"))
        trides.Mode = CipherMode.ECB
        Return System.Text.ASCIIEncoding.ASCII.GetString(trides.CreateDecryptor.TransformFinalBlock(enc, 0, enc.Length))
    End Function

    'Private Function createPRoduct(brand As String, category As String, code As String, description As String, unit As String, qty As String, unit_price As String)

    '    rs.Open("select * from tblproducts", connection, 2, 2)
    '    rs.AddNew()
    '    rs("quantity").Value = qty
    '    rs("category").Value = category
    '    rs("brand").Value = brand
    '    rs("code").Value = code

    '    rs("description").Value = description
    '    rs("unit").Value = unit
    '    rs("unit_price").Value = unit_price
    '    rs.Update()
    '    rs.Close()
    '    Return Nothing
    'End Function
    'Public Function getAllproducts()

    '    rs1.Open("select * from tblproduct", connection1, 2, 2)
    '    Do Until rs1.EOF
    '        createPRoduct(rs1("brand").Value, rs1("category").Value, rs1("code").Value, rs1("description").Value, rs1("unit").Value, rs1("quantity").Value, rs1("unitprice").Value)
    '        rs1.MoveNext()
    '    Loop
    '    rs1.Close()

    'End Function

End Module
