<%@language=Vbscript%>
<%option explicit%>
<html>
<head>
  <title>Search</title>
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

<div class="wrapper">
	<h2>&nbsp&nbsp&nbspNo of visits by all</h2>
	<div class="content" style="display:inline-block;">
		There have been <%=Application("count")%> visit(s) on our website including everyone.
	</div>
	<h2>&nbsp&nbsp&nbspNo of visits using this device & browser</h2>
	<div class="content" style="display:inline-block;">
		There have been <%=request.cookies("count")%> visit(s) on our website using this device & browser.
	</div>
	
</div>
</body>
</html>