<%@language=Vbscript%>
<%option explicit%>
<%
	dim id
	id=session("id")
%>
<html>
<head>
  <title>Search</title>
  <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<%
dim con,res
set con = Server.CreateObject("ADODB.Connection")
	con.Provider = "Microsoft.Jet.OLEDB.4.0"
	con.Open "C:/inetpub/wwwroot/project/project.mdb"

	set res = Server.CreateObject("ADODB.RecordSet")
	res.Open "cart",con,0,3,2
	res.AddNew
	res("ID") =id
	res("username")=session("username")
	res.Update
	res.close
	con.close
	response.redirect("index.html")
%>
</body>
</html>