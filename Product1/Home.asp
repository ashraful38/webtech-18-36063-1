<!DOCTYPE html>
<html>
<head>
 <link rel="stylesheet" href="css/style.css">
  <style>
    .display{
	 margin-left:500px;

	}
	.display table{
		border:1px solid blue;
		padding:4px;
		text-align:center;
		margin-left:10px;
		border-collapse: collapse;
	}
	
	.display td{
		border:1px solid blue;
		padding:4px;
		text-align:center;
		margin-left:10px;
		border-collapse: collapse;
	}	
	
</style>
</head>
<body>
 <nav class="header">
    <h1 class="logo">Inventory Management</h1>
    <div>
	   	<ul class="ul-item">
			<li><a class="active" href="Home.asp">Home</a></li>
			<li><a href="Product.asp">Product</a></li>
			<li><a href="Supplier.asp">Supplier</a></li>
			<li><a href="purchase.asp">Purchase</a></li>
	    </ul>
	</div>
	<div>
	   	<ul class="ul-right-item">
			<li><a href="#contact">login</a></li>
			<li><a href="#about">Registration</a></li>
	    </ul>
	</div>
 </nav>
 
<h3 style="color:blue;text-align:center;">Our Product</h3>
  
 <div class="display">
 
     <form  action="Supplier.asp">
		 <table style="text-align:center;">
			 <tr>
				<td>ProductId</td>
				<td>ProductName</td>
				<td>Catagory</td>
			</tr>
			<% dim db_connection
				db_connection = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"

				set conn = Server.CreateObject("ADODB.Connection")
				conn.open(db_connection)
				set rs=conn.execute("viewproductlist")
                dim x
				do until rs.EOF
					Response.Write("<tr>")
					  for each X in rs.Fields
						 Response.Write("<td>" & x.value & "</td>")
					  Next
					  
					Response.write("</tr>")
					rs.movenext
				loop

			%>

		</table>
	</form>
 </div>


</body>
</html>