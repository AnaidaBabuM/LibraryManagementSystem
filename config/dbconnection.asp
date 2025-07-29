<%
Function GetDBConnection()
    Dim conn, connStr
    Set conn = Server.CreateObject("ADODB.Connection")
    connStr = "Provider=SQLOLEDB;Data Source=" & Server.CreateObject("WScript.Shell").Environment("Process")("DB_SERVER") & ";" & _
              "Initial Catalog=" & Server.CreateObject("WScript.Shell").Environment("Process")("DB_NAME") & ";" & _
              "User ID=" & Server.CreateObject("WScript.Shell").Environment("Process")("DB_USER") & ";" & _
              "Password=" & Server.CreateObject("WScript.Shell").Environment("Process")("DB_PASSWORD") & ";"
    conn.Open connStr
    Set GetDBConnection = conn
End Function
%>
