<%
Class Book
    Public Id, Title, Author, ISBN, PublishedYear

    Public Function GetAllBooks()
        Dim conn, rs, books
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Books", conn, 3, 3
        Set books = CreateObject("Scripting.Dictionary")
        While Not rs.EOF
            Dim book
            Set book = New Book
            book.Id = rs("Id")
            book.Title = rs("Title")
            book.Author = rs("Author")
            book.ISBN = rs("ISBN")
            book.PublishedYear = rs("PublishedYear")
            books.Add CStr(book.Id), book
            rs.MoveNext
        Wend
        rs.Close
        conn.Close
        Set GetAllBooks = books
    End Function

    Public Function GetBookById(bookId)
        Dim conn, rs
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Books WHERE Id = " & bookId, conn, 3, 3
        If Not rs.EOF Then
            Id = rs("Id")
            Title = rs("Title")
            Author = rs("Author")
            ISBN = rs("ISBN")
            PublishedYear = rs("PublishedYear")
        End If
        rs.Close
        conn.Close
    End Function

    Public Sub Save()
        Dim conn, sql
        Set conn = GetDBConnection()
        If Id = 0 Then
            sql = "INSERT INTO Books (Title, Author, ISBN, PublishedYear) VALUES ('" & Replace(Title, "'", "''") & "', '" & Replace(Author, "'", "''") & "', '" & Replace(ISBN, "'", "''") & "', " & PublishedYear & ")"
        Else
            sql = "UPDATE Books SET Title = '" & Replace(Title, "'", "''") & "', Author = '" & Replace(Author, "'", "''") & "', ISBN = '" & Replace(ISBN, "'", "''") & "', PublishedYear = " & PublishedYear & " WHERE Id = " & Id
        End If
        conn.Execute sql
        conn.Close
    End Sub
End Class
%>
