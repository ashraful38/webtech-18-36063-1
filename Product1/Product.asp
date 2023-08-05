<%      
    'Get the value of parameter from querystring     
    Catagory_ID = Request.Querystring("CatagoryId")    
    
    'Declare a variable to store the connection string    
    Dim connstr    
    connstr = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"   
    
    'Create a new ADODB recordset      
    set rsSelectedProduct = Server.CreateObject("ADODB.recordset")     
    
    'Create a new ADODB connection     
    Set conn = Server.CreateObject("ADODB.Connection")     
    
    'Open the connection using the connection string      
    conn.open connstr    
    
    'Store  the query in a variable , in real world scenario we should be using a stored procedure with parameters    
    QuerySQL = "Select  CatagoryName from Catagory where CatagoryId = '" & Catagory_ID & "'"     
    
    'Execute the query     
    set rsSelectedProduct = conn.Execute(QuerySQL)      
    
    'Check if recordset is empty , if not then store the value     
    if not rsSelectedProduct.EOF then      
      product = rsSelectedProduct("CatagoryName")      
    else      
      product = ""      
    end if    
         
    'Close the Connection and recordset    
    
    rsSelectedProduct.Close    
    conn.Close    
    Set rsSelectedProduct = Nothing    
    Set conn = Nothing   
  
    'Declare a function to set the selected text     
    Function selectProduct(vProduct)      
      if vProduct = product then      
          Response.Write("selected=""selected""")           
     end if      
    End Function    
  
    'Create Connection object    
    Set conn2 = Server.CreateObject("ADODB.Connection")      
    'Create Recordset object    
    Set rsProductList = Server.CreateObject("ADODB.recordset")    
    
    'Open the connection using the previous connection string  
    conn2.open connstr      
    
     'Declare a variable to store the query to be excuted. Here we are using a hardcoded query.We can also use a stored procedure    
    QueryProduct = "select CatagoryName,CatagoryId from Catagory order by CatagoryName asc"    
  
     'Execute the query   
    set rsProductList = conn2.Execute(QueryProduct)    
     'Declare the array  
    Dim arrProducts   
    
    if not rsProductList.EOF then    
          arrProducts = rsProductList.GetRows()  ' Convert recordset to 2D Array  
    end if   
      
   'Close the Connection and recordset    
    
    rsProductList.Close    
    conn2.Close    
    Set rsProductList = Nothing    
    Set conn2 = Nothing   
%>   

  <%
	    dim db_connection
			db_connection = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"

			set conn = Server.CreateObject("ADODB.Connection")
			conn.open(db_connection)
	   If Request.Form("submit") <> "" then
			
			


                
                ProductName = Trim(request.form("ProductName"))
			    productSelect = Trim(request.form("productSelect"))
				
				conn.execute  "INSERT INTO product VALUES('" & ProductName & "', '" & productSelect& "')"
				 If Err.Number <> 0 Then
					Response.Write("Error: " & Err.Description)
				Else
					'Response.Write("Data has been submitted.")
				End If
			     

				'On Error Resume Next
				'conn.execute(ssSQL)
		
			'conn.Close()
			'set conn = Nothing
		 END IF
		 set rs=conn.execute("select * from product")
%>


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
    <h3 style="color:blue;text-align:center;">Product view</h3>
 </div>


<div class=input-area>
    <form method="post" action="">
		<table>
			<tr>
				<td>ProductName:</td>
				<td><input name="ProductName"></td>
			</tr>
			<tr>
				<td>Catagory:</td>
				<td>    
                   
				   <select name="productSelect" class="dropdowitem" id="productSelect">    
					  <%    
						  'Check whether it's a proper array or not    
							if IsArray(arrProducts) then    
						
								For i = 0 to ubound(arrProducts, 2) %>    
						
								<option value="<%= arrProducts(1,i)%>" <%= selectProduct(arrProducts(0,i)) %>> <%= arrProducts(0,i) %> </option>    
						
							   <% next %>    
						
							<% else %>    
						
								<option value=""> Select </option>    
						
							<% end if %>    
						
						</select>  
			
				</td>
			</tr>
		</table>
		<br><br>
		<input type="submit" name="submit" value="Add New">
		<input type="reset" value="Cancel">
    </form>

 </div>
 
 <br>
<h3>display data table</h3>
</br>

 
  <div class="display">
 
     <form  action="product.asp">
		 <table>
			 <tr>
				<td>ProductId</td>
				<td>ProductName</td>
				<td>CatagoryId</td>
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

