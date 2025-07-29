<!-- #include file="../config/dbconnection.asp" -->
<!-- #include file="../models/loan.asp" -->
<!-- #include file="../models/book.asp" -->
<!-- #include file="../models/member.asp" -->
<%
Function GetLoans()
    Dim loan
    Set loan = New Loan
    Set GetLoans = loan.GetAllLoans()
End Function

Function GetLoan(id)
    Dim loan
    Set loan = New Loan
    loan.GetLoanById(id)
    Set GetLoan = loan
End Function

Sub BorrowBook(bookId, memberId)
    Dim loan
    Set loan = New Loan
    loan.BookId = bookId
    loan.MemberId = memberId
    loan.BorrowDate = Date()
    loan.Save
End Sub

Sub ReturnLoan(id)
    Dim loan
    Set loan = New Loan
    loan.Id = id
    loan.ReturnBook
End Sub
%>
