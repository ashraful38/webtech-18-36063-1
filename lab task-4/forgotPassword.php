<!DOCTYPE html>
<html>
<head>
	<title>Public Home</title>
</head>
<body>
<table border="1" width="100%">
	<tr>
		<td align="right">
			 <a href="public_home.php"> <img src="download.png" align="left"> </a>
			 <a href="public_home.php"> Home </a> |
			<a href="login.php"> Login </a> |
			<a href="registration.php"> Registration </a> 
			
		</td>
	</tr>
		<tr height="300px">
 		<td colspan="2" align="center">
		<br />  
           <div class="container" style="width:500px;">  
                                
                <form method="post">  
                     <?php   
                     if(isset($error))  
                     {  
                          echo $error;  
                     }  
                     ?>  
                     <br /> 
                      <fieldset> 
                     <legend>FORGET PASSWORD</legend> 
                     
                     <label>Enter Email</label>
                     <input type="text" name = "un" class="form-control" /><br />
                     

 


                   <hr>
                     
                     <input type="submit" name="submit" value="Submit" class="btn btn-info" /> 
                                        
                     <?php  
                     if(isset($message))  
                     {  
                          echo $message;  
                     }  
                     ?>  
                </form>  
           </div>  
           <br />  


		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="2" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
</body>
</html>

      