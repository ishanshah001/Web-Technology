<%@language=Vbscript%>
<%option explicit%>
<%
	dim id,username
	id=session("id")
	username=session("username")
%>
<html>
<head>
  <title>Search</title>
  <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<%
dim con,res,cmd,sql
set con = Server.CreateObject("ADODB.Connection")
	con.Provider = "Microsoft.Jet.OLEDB.4.0"
	con.Open "C:/inetpub/wwwroot/project/project.mdb"

	set res = Server.CreateObject("ADODB.RecordSet")
	sql = "INSERT INTO [order] VALUES("&id&",'"&username&"')"
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = con
	cmd.CommandText = sql
	set res = cmd.Execute
	set res = nothing
	con.close
	set con = nothing
	response.redirect("cartpage.asp")
%>
</body>
</html>