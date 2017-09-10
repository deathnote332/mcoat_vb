Module LoginController

    Public Function authenticate(username As String, password As String) As Boolean

        rs.Open("select * from users where username='" & username & "' and password='" & encrypt(password) & "'", connection, 2, 2)
        If Not rs.EOF Then
            Dim userType As String = rs("user_type").Value
            Dim fullname As String = rs("first_name").Value & " " & rs("last_name").Value
            If userType = "admin" Then
                MsgBox("Access granted!" & vbCrLf & "Welcome " & fullname & vbCrLf & vbCrLf & "You logged on as an administrator", vbInformation, "System")
            Else
                MsgBox("Access granted!" & vbCrLf & "Welcome " & fullname, vbInformation, "System")
            End If
            rs.Close()
            Return True
        Else
            MsgBox("Invalid credentials", vbExclamation, "Error")
            rs.Close()
            Return False
        End If

        rs.Close()
        Return Nothing

    End Function





End Module
