<%
Sub GetBooks(page)
    Dim rs, sql
    OpenConnection
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM Books ORDER BY Title OFFSET " & ((page-1) * ITEMS_PER_PAGE) & " ROWS FETCH NEXT " & ITEMS_PER_PAGE & " ROWS ONLY"
    rs.Open sql, Conn, 1, 1
    If Not rs.EOF Then
        Response.Write "<table><tr><th>Title</th><th>Author</th><th>ISBN</th><th>Actions</th></tr>"
        Do While Not rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("Title") & "</td>"
            Response.Write "<td>" & rs("Author") & "</td>"
            Response.Write "<td>" & rs("ISBN") & "</td>"
            Response.Write "<td><a href='edit.asp?id=" & rs("BookID") & "'>Edit</a> | <a href='delete.asp?id=" & rs("BookID") & "'>Delete</a></td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop
        Response.Write "</table>"
    End If
    rs.Close
    Set rs = Nothing
    CloseConnection
End Sub

Sub AddBook(title, author, isbn)
    OpenConnection
    Dim sql
    sql = "INSERT INTO Books (Title, Author, ISBN) VALUES ('" & Replace(title, "'", "''") & "', '" & Replace(author, "'", "''") & "', '" & Replace(isbn, "'", "''") & "')"
    Conn.Execute sql
    CloseConnection
End Sub
%>
