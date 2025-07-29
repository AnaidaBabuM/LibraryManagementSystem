<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/loanService.asp" -->
<!-- #include file="../../services/bookService.asp" -->
<!-- #include file="../../services/memberService.asp" -->
<%
Dim books, members
Set books = GetBooks()
Set members = GetMembers()
If Request.Form("submit") <> "" Then
    Call BorrowBook(Request.Form("bookId"), Request.Form("memberId"))
    Response.Redirect "list.asp"
End If
%>
<h2>Borrow Book</h2>
<form method="post">
    <label>Book:</label>
    <select name="bookId">
        <%
        For Each bookId In books
            Dim book
            Set book = books(bookId)
        %>
            <option value="<%=book.Id%>"><%=book.Title%></option>
        <%
        Next
        %>
    </select><br>
    <label>Member:</label>
    <select name="memberId">
        <%
        For Each memberId In members
            Dim member
            Set member = members(memberId)
        %>
            <option value="<%=member.Id%>"><%=member.FirstName & " " & member.LastName%></option>
        <%
        Next
        %>
    </select><br>
    <input type="submit" name="submit" value="Borrow">
</form>
<!-- #include file="../../includes/footer.asp" -->
