<%      
    'Get the value of parameter from querystring     
    Product_ID = Request.Querystring("ProductId")    
    
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
    QuerySQL = "Select  ProductName from product where ProductId = '" & Product_ID & "'"     
    
    'Execute the query     
    set rsSelectedProduct = conn.Execute(QuerySQL)      
    
    'Check if recordset is empty , if not then store the value     
    if not rsSelectedProduct.EOF then      
      product = rsSelectedProduct("ProductName")      
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
    QueryProduct = "select ProductName,ProductId from product order by ProductName asc"    
  
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
    'Get the value of parameter from querystring     
    catagory_ID = Request.Querystring("SupplierId")    
    
    'Declare a variable to store the connection string    
    Dim connstr1    
    connstr1 = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"   
    
    'Create a new ADODB recordset      
    set rsSelectedProduct = Server.CreateObject("ADODB.recordset")     
    
    'Create a new ADODB connection     
    Set conn = Server.CreateObject("ADODB.Connection")     
    
    'Open the connection using the connection string      
    conn.open connstr    
    
    'Store  the query in a variable , in real world scenario we should be using a stored procedure with parameters    
    QuerySQL = "Select  SupplierName from Suppliers where SupplierId = '" & catagory_ID & "'"     
    
    'Execute the query     
    set rsSelectedcatagory = conn.Execute(QuerySQL)      
    
    'Check if recordset is empty , if not then store the value     
    if not rsSelectedcatagory.EOF then      
      Catagory = rsSelectedcatagory("SupplierName")      
    else      
      Catagory = ""      
    end if    
         
    'Close the Connection and recordset    
    
    rsSelectedcatagory.Close    
    conn.Close    
    Set rrsSelectedcatagory = Nothing    
    Set conn = Nothing   
  
    'Declare a function to set the selected text     
    Function selectCatagory(vcatagory)      
      if vcatagory = Catagory then      
          Response.Write("selected=""selected""")           
     end if      
    End Function    
  
    'Create Connection object    
    Set conn2 = Server.CreateObject("ADODB.Connection")      
    'Create Recordset object    
    Set rsCatagorytList = Server.CreateObject("ADODB.recordset")    
    
    'Open the connection using the previous connection string  
    conn2.open connstr1     
    
     'Declare a variable to store the query to be excuted. Here we are using a hardcoded query.We can also use a stored procedure    
    QueryProduct = "select SupplierName,  SupplierId from Suppliers order by SupplierName asc"    
  
     'Execute the query   
    set rsCatagorytList = conn2.Execute(QueryProduct)    
     'Declare the array  
    Dim arrProducts1   
    
    if not rsCatagorytList.EOF then    
          arrProducts1 = rsCatagorytList.GetRows()  ' Convert recordset to 2D Array  
    end if   
      
   'Close the Connection and recordset    
    
    rsCatagorytList.Close    
    conn2.Close    
    Set rsCatagorytList = Nothing    
    Set conn2 = Nothing   
%>  



  <%
	    dim db_connection
			db_connection = "Provider=MSDASQL; DRIVER={SQL Server}; SERVER=ASHRAFUL_MIS; DATABASE=product1; Trusted_Connection=yes;"

			set conn = Server.CreateObject("ADODB.Connection")
			conn.open(db_connection)
	   If Request.Form("submit") <> "" then
			
			


                
                ProductName = Trim(request.form("productSelect"))
			    SupplierName = Trim(request.form("SupplierSelect"))
			    color = Trim(request.form("color"))
			    Quantity = Trim(request.form("quantity"))
				
				conn.execute  "INSERT INTO Purchase VALUES('" & ProductName & "', '" & SupplierName& "' , '" &  color& "' , '" &  Quantity& "')"
				 If Err.Number <> 0 Then
					'Response.Write("Error: " & Err.Description)
				Else
					'Response.Write("Data has been submitted.")
				End If
			     

				'On Error Resume Next
				'conn.execute(ssSQL)
		
			'conn.Close()
			'set conn = Nothing
		 END IF
		 set rs=conn.execute("ViewPurchaseGrid1")
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
    <h3 style="color:blue;text-align:center;">Purchase view</h3>
 </div>
 
 <div class=input-area>
    <form method="post" action="">
		<table>
			<tr>
				<td>ProductName:</td>
				<td>
				   <select name="productSelect" class="dropdowitem" id="productSelect">    
					  <%    
						  'Check whether it's a proper array or not    
							if IsArray(arrProducts) then    
						
								For i = 0 to ubound(arrProducts, 2) %>    
						
								<option value="<%= arrProducts(1,i)%>" <%= selectCatagory(arrProducts(0,i)) %>> <%= arrProducts(0,i) %> </option>    
						
							   <% next %>    
						
							<% else %>    
						
								<option value=""> Select </option>    
						
							<% end if %>    
						
						</select> 
				 </td>
			</tr>
			<tr>
				<td>SupplierName:</td>
				<td>
				  <select name="SupplierSelect" class="dropdowitem" id="SupplierSelect">    
					  <%    
						  'Check whether it's a proper array or not    
							if IsArray(arrProducts) then    
						
								For i = 0 to ubound(arrProducts1, 2) %>    
						
								<option value="<%= arrProducts1(1,i)%>" <%= selectProduct(arrProducts1(0,i)) %>> <%= arrProducts1(0,i) %> </option>    
						
							   <% next %>    
						
							<% else %>    
						
								<option value=""> Select </option>    
						
							<% end if %>    
						
						</select> 
			
				</td>
			</tr>
			<tr>
				<td>Color:</td>
				<td><input name="color"></td>
			</tr>
			<tr>
				<td>Quantity:</td>
				<td><input name="quantity"></td>
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
				<td>PurchaseId</td>
				<td>ProductName</td>
				<td>SupplierName</td>
				<td>Color</td>
				<td>Quantity</td>
				
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


