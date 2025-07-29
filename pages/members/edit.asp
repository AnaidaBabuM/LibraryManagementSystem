<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/memberService.asp" -->
<%
Dim id, member
id = Request.QueryString("id")
Set member = GetMember(id)
If Request.Form("submit") <> "" Then
    Call SaveMember(id, Request.Form("firstName"), Request.Form("lastName"), Request.Form("email"), Request.Form("joinDate"))
    Response.Redirect "list.asp"
End If
%>
<h2>Edit Member</h2>
<form method="post">
    <label>First Name:</label><input type="text" name="firstName" value="<%=member.FirstName%>"><br>
    <label>Last Name:</label><input type="text" name="lastName" value="<%=member.LastName%>"><br>
    <label>Email:</label><input type="text" name="email" value="<%=member.Email%>"><br>
    <label>Join Date:</label><input type="text" name="joinDate" value="<%=member.JoinDate%>"><br>
    <input type="submit" name="submit" value="Save">
</form>
<!-- #include file="../../includes/footer.asp" -->
