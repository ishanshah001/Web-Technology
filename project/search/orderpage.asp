<%@language=Vbscript%>
<%option explicit%>
<%
	dim username
	username=session("username")
%>

<html>
<head>
  <titleOrdered</title>
  <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
	<div class="menu">
			<div class="logo"><img src ="logo.jpg"></div>
			<ul>
				<li><a href="http://localhost/project/search/visits.asp"><img src="images/5.jpg"></a></li>
				<li><a href="http://localhost/project/search/orderpage.asp"><img src="images/6.svg"></a></li>
				<li><a href="http://localhost/project/search/cartpage.asp"><img src="images/7.png"></a></li>
			</ul>
	</div>
	<div class="sidebar">
		<a href="http://localhost/project/dashboard/index.html"><span>Dashboard</span></a>
		<a href="http://localhost/project/search/index.html"><span>Search</span></a>
		<a href="http://localhost/project/dashboard/index.html"><span>Recommend</span></a>
		<a href="http://localhost/project/search/cartpage.asp"><span>Cart</span></a>
		<a href="http://localhost/project/search/orderpage.asp"><span>Your orders</span></a>
		<a href="http://localhost/project/search/upcoming.html"><span>Upcoming</span></a>
	 </div> 
<%
dim con,res,sql
set con = Server.CreateObject("ADODB.Connection")
	con.Provider = "Microsoft.Jet.OLEDB.4.0"
	con.Open "C:/inetpub/wwwroot/project/project.mdb"

	set res = Server.CreateObject("ADODB.RecordSet")
	sql="SELECT * FROM [order] INNER JOIN [diamond] ON order.ID = diamond.ID where order.username='"&username&"' ORDER BY order.ID"
	res.open sql,con
%>

<div class="wrapper">
	<h2>&nbsp&nbsp&nbspItems you have ordered</h2>
	<div class="content" style="display:inline-block;">
	<table cellspacing="10" cellpadding="10">
	<%
	Response.write("<tr><th>ID</th>")
	Response.write("<th>Shape</th>")
	Response.write("<th>Carat</th>")
	Response.write("<th>Color Desc</th>")
	Response.write("<th>Color</th>")
	Response.write("<th>Clarity</th>")
	Response.write("<th>Fluorescence</th>")
	Response.write("<th>Shade</th>")
	Response.write("<th>Lab</th>")
	Response.write("<th>Cut</th>")
	Response.write("<th>Polish</th>")
	Response.write("<th>Symmetry</th>")
	Response.write("<th>Location</th></tr>")

do while not res.EOF
	Response.write("<tr><td>"&res("order.ID")&"</td>")
	Response.write("<td>"&res("shape")&"</td>")
	Response.write("<td>"&res("carat")&"</td>")
	Response.write("<td>"&res("color_desc")&"</td>")
	Response.write("<td>"&res("color")&"</td>")
	Response.write("<td>"&res("clarity")&"</td>")
	Response.write("<td>"&res("Fluorescence")&"</td>")
	Response.write("<td>"&res("shade")&"</td>")
	Response.write("<td>"&res("lab")&"</td>")
	Response.write("<td>"&res("cut")&"</td>")
	Response.write("<td>"&res("polish")&"</td>")
	Response.write("<td>"&res("symmetry")&"</td>")
	Response.write("<td>"&res("location")&"</td></tr>")

	res.MoveNext
Loop
res.close
con.close
%></table>
	</div>
</div>
</body>
</html>