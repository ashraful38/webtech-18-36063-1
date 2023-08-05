<!DOCTYPE html>
<html>
<head>
 
 <link rel="stylesheet" href="css/style.css">
 <style>
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
 <nav class="header cheader">
    <h1 class="logo">Inventory Management</h1>
    <div>
	   	<ul style="text-align:center"  class="ul-item">
			<li><a class="active" href="Home.asp">Home</a></li>
			<li><a href="Product.asp">Product</a></li>
			<li><a href="Supplier.asp">Supplier</a></li>
			<li><a href="purchase.asp">Purchase</a></li>
	    </ul>
	</div>
	
 </nav>
 
 <div>
    <h3 style="color:blue;text-align:center;">Supplier view</h3>
 </div>


<div class=input-area>
    <form action="" method="post" >
		<table>
			<tr>
				<td>SupplierName:</td>
				<td><input name="SupplierName"></td>
			</tr>
			<tr>
			    <td>SupplierAddress:</td>
				<td><input name="SupplierAddress"></td>
			</tr>
		</table>
		<br><br>
		<input type="submit" name="submit"  value="Add New">
		<input type="reset"  value="Cancel">
    </form>

 </div>
 


  <%
	    dim db_connection
			db_connection = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"

			set conn = Server.CreateObject("ADODB.Connection")
			conn.open(db_connection)
	   If Request.Form("submit") <> "" then
			
			


                
                SupplierName = Trim(request.form("SupplierName"))
			    SupplierAddress = Trim(request.form("SupplierAddress"))
				
				conn.execute  "INSERT INTO Suppliers VALUES('" & SupplierName & "', '" & SupplierAddress & "')"
				 If Err.Number <> 0 Then
					Response.Write("Error: " & Err.Description)
				Else
					Response.Write("Data has been submitted.")
				End If
			     

				'On Error Resume Next
				'conn.execute(ssSQL)
		
			'conn.Close()
			'set conn = Nothing
		 END IF	
		set rs=conn.execute("select * from Suppliers")
%>

<br>
<h3>display data table</h3>
</br>




 
 <div class="display">
 
     <form  action="Supplier.asp">
		 <table>
			 <tr>
				<td>SupplierId</td>
				<td>SupplierName</td>
				<td>SupplierAddress</td>
			</tr>
			<%
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