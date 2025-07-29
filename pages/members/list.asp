<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/memberService.asp" -->
<%
Dim members
Set members = GetMembers()
%>
<h2>Members</h2>
<a href="add.asp">Add New Member</a>
<table border="1">
    <tr>
        <th>First Name</th>
        <th>Last Name</th>
        <th>Email</th>
        <th>Join Date</th>
        <th>Actions</th>
    </tr>
    <%
    For Each memberId In members
        Dim member
        Set member = members(memberId)
    %>
    <tr>
        <td><%=member.FirstName%></td>
        <td><%=member.LastName%></td>
        <td><%=member.Email%></td>
        <td><%=member.JoinDate%></td>
        <td><a href="edit.asp?id=<%=member.Id%>">Edit</a> | <a href="list.asp?delete=<%=member.Id%>" onclick="return confirm('Are you sure?')">Delete</a></td>
    </tr>
    <%
    Next
    %>
</table>
<% If Request.QueryString("delete") <> "" Then
    Call DeleteMember(Request.QueryString("delete"))
    Response.Redirect "list.asp"
End If %>
<!-- #include file="../../includes/footer.asp" -->
