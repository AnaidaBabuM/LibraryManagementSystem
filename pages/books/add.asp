<!-- #include file="../../includes/config.asp" -->
<!-- #include file="../../includes/db_connect.asp" -->
<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/book_service.asp" -->
<h2>Add New Book</h2>
<%
If Request.Form("submit") <> "" Then
    Call AddBook(Request.Form("title"), Request.Form("author"), Request.Form("isbn"))
    Response.Redirect "list.asp"
End If
%>
<form method="post">
    <div class="form-group">
        <label>Title:</label>
        <input type="text" name="title" required>
    </div>
    <div class="form-group">
        <label>Author:</label>
        <input type="text" name="author" required>
    </div>
    <div class="form-group">
        <label>ISBN:</label>
        <input type="text" name="isbn" required>
    </div>
    <input type="submit" name="submit" value="Add Book">
</form>
</div>
</body>
</html>
