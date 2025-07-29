<%
Sub OpenConnection()
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open DB_CONNECTION_STRING
End Sub

Sub CloseConnection()
    If IsObject(Conn) Then
        Conn.Close
        Set Conn = Nothing
    End If
End Sub
%>
