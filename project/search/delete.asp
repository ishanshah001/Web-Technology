<%@language=Vbscript%>
<%option explicit%>
<%
	dim id, username
	id=request.form("id")
	username=session("username")
%>
<html>
<body>
<% 
dim conn,rs,sql,cmd
set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "C:/inetpub/wwwroot/project/project.mdb"
sql = "delete from cart where(ID="&id&" AND username='"&username&"')"
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = sql
set rs = cmd.Execute
set rs = nothing
conn.close
set conn = nothing
response.redirect("cartpage.asp")
%>
</body>
</html>