<%@language=Vbscript%>
<%option explicit%>
<%
	dim shape,from_,to_,color_d,c,cl,f,s,l,cu,p,sm,lo
	shape=request.form("shape")
	<!--from_=request.form("from")
	<!--to_=request.form("to")
	color_d=request.form("t")
	c=request.form("c")
	cl=request.form("cl")
	f=request.form("f")
	s=request.form("s")
	l=request.form("l")
	p=request.form("p")
	cu=request.form("cu")
	sm=request.form("sm")
	lo=request.form("lo")
%>
<html>
<head>
  <title>Search</title>
  <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
	<div class="menu">
			<div class="logo"><img src ="logo.jpg"></div>
			<ul>
				<li><a href="#"><img src="images/5.svg"></a></li>
				<li><a href="orderpage.asp"><img src="images/6.svg"></a></li>
				<li><a href="cartpage.asp"><img src="images/7.png"></a></li>
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
dim conn,rs
set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "C:/inetpub/wwwroot/project/project.mdb"

set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "select * from diamond ORDER BY id",conn
%>
<div class="wrapper">
<h2>&nbsp&nbsp&nbspSearched Diamonds</h2>
<div class="content">
<table cellspacing="10" cellpadding="5">
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

do while not rs.EOF
	Response.write("<tr><td>"&rs("id")&"</td>")
	Response.write("<td>"&rs("shape")&"</td>")
	Response.write("<td>"&rs("carat")&"</td>")
	Response.write("<td>"&rs("color_desc")&"</td>")
	Response.write("<td>"&rs("color")&"</td>")
	Response.write("<td>"&rs("clarity")&"</td>")
	Response.write("<td>"&rs("Fluorescence")&"</td>")
	Response.write("<td>"&rs("shade")&"</td>")
	Response.write("<td>"&rs("lab")&"</td>")
	Response.write("<td>"&rs("cut")&"</td>")
	Response.write("<td>"&rs("polish")&"</td>")
	Response.write("<td>"&rs("symmetry")&"</td>")
	Response.write("<td>"&rs("location")&"</td></tr>")

	rs.MoveNext
Loop
rs.close
conn.close
%>
</table>
</div>
<h2>&nbsp&nbsp&nbspAdd to cart:</h2>
<div class="content">
<form method="POST" action="cart.asp">
Enter id no:<input type="number" name="id" style="z-index:10;"><br><br><br>
<input type="submit" value="Submit">
</form>
</div>
</body>
</html>

