<%
Class Loan
    Public Id, BookId, MemberId, BorrowDate, ReturnDate

    Public Function GetAllLoans()
        Dim conn, rs, loans
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT l.*, b.Title, m.FirstName, m.LastName FROM Loans l JOIN Books b ON l.BookId = b.Id JOIN Members m ON l.MemberId = m.Id", conn, 3, 3
        Set loans = CreateObject("Scripting.Dictionary")
        While Not rs.EOF
            Dim loan
            Set loan = New Loan
            loan.Id = rs("Id")
            loan.BookId = rs("BookId")
            loan.MemberId = rs("MemberId")
            loan.BorrowDate = rs("BorrowDate")
            loan.ReturnDate = rs("ReturnDate")
            loans.Add CStr(loan.Id), loan
            rs.MoveNext
        Wend
        rs.Close
        conn.Close
        Set GetAllLoans = loans
    End Function

    Public Function GetLoanById(loanId)
        Dim conn, rs
        Set conn = GetDBConnection()
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM Loans WHERE Id = " & loanId, conn, 3, 3
        If Not rs.EOF Then
            Id = rs("Id")
            BookId = rs("BookId")
            MemberId = rs("MemberId")
            BorrowDate = rs("BorrowDate")
            ReturnDate = rs("ReturnDate")
        End If
        rs.Close
        conn.Close
    End Function

    Public Sub Save()
        Dim conn, sql
        Set conn = GetDBConnection()
        If Id = 0 Then
            sql = "INSERT INTO Loans (BookId, MemberId, BorrowDate, ReturnDate) VALUES (" & BookId & ", " & MemberId & ", '" & BorrowDate & "', NULL)"
        Else
            sql = "UPDATE Loans SET BookId = " & BookId & ", MemberId = " & MemberId & ", BorrowDate = '" & BorrowDate & "', ReturnDate = " & IIf(IsEmpty(ReturnDate), "NULL", "'" & ReturnDate & "'") & " WHERE Id = " & Id
        End If
        conn.Execute sql
        conn.Close
    End Sub

    Public Sub ReturnBook()
        Dim conn
        Set conn = GetDBConnection()
        conn.Execute "UPDATE Loans SET ReturnDate = GETDATE() WHERE Id = " & Id
        conn.Close
    End Sub
End Class
%>
