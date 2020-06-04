<%@language=Vbscript%>
<%option explicit%>
<%
	dim id
	id=request.form("id")
	session("id")=id
%>
<html>
<head>
  <title>Search</title>
  <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<%
dim conn,rs,con,res
set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "C:/inetpub/wwwroot/project/project.mdb"

set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "select * from diamond where id="&id&"",conn
%>
<%
if rs.EOF=True then
	response.redirect("wrongcart.html")
else
	response.redirect("addorder.asp")
end if
rs.close
conn.close
%>
</body>
</html>