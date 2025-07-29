<!-- #include file="../../includes/config.asp" -->
<!-- #include file="../../includes/db_connect.asp" -->
<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/book_service.asp" -->
<h2>Books List</h2>
<a href="add.asp">Add New Book</a>
<%
Dim page
page = CInt(Request.QueryString("page"))
If page <= 0 Then page = 1
Call GetBooks(page)
%>
</div>
</body>
</html>
