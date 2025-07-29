<!-- #include file="../../includes/header.asp" -->
<!-- #include file="../../services/loanService.asp" -->
<%
If Request.QueryString("id") <> "" Then
    Call ReturnLoan(Request.QueryString("id"))
    Response.Redirect "list.asp"
End If
%>
<!-- #include file="../../includes/footer.asp" -->
