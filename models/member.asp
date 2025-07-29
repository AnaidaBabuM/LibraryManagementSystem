<%
Class Member
    Public Id, FirstName, LastName, Email, JoinDate

    Public Function GetAllMembers()
        Dim conn, rs, members
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Members", conn, 3, 3
        Set members = CreateObject("Scripting.Dictionary")
        While Not rs.EOF
            Dim member
            Set member = New Member
            member.Id = rs("Id")
            member.FirstName = rs("FirstName")
            member.LastName = rs("LastName")
            member.Email = rs("Email")
            member.JoinDate = rs("JoinDate")
            members.Add CStr(member.Id), member
            rs.MoveNext
        Wend
        rs.Close
        conn.Close
        Set GetAllMembers = members
    End Function

    Public Function GetMemberById(memberId)
        Dim conn, rs
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Members WHERE Id = " & memberId, conn, 3, 3
        If Not rs.EOF Then
            Id = rs("Id")
            FirstName = rs("FirstName")
            LastName = rs("LastName")
            Email = rs("Email")
            JoinDate = rs("JoinDate")
        End If
        rs.Close
        conn.Close
    End Function

    Public Sub Save()
        Dim conn, sql
        Set conn = GetDBConnection()
        If Id = 0 Then
            sql = "INSERT INTO Members (FirstName, LastName, Email, JoinDate) VALUES ('" & Replace(FirstName, "'", "''") & "', '" & Replace(LastName, "'", "''") & "', '" & Replace(Email, "'", "''") & "', '" & JoinDate & "')"
        Else
            sql = "UPDATE Members SET FirstName = '" & Replace(FirstName, "'", "''") & "', LastName = '" & Replace(LastName, "'", "''") & "', Email = '" & Replace(Email, "'", "''") & "', JoinDate = '" & JoinDate & "' WHERE Id = " & Id
        End If
        conn.Execute sql
        conn.Close
    End Sub

    Public Sub Delete()
        Dim conn
        Set conn = GetDBConnection()
        conn.Execute "DELETE FROM Members WHERE Id = " & Id
        conn.Close
    End Sub
End Class
%>
