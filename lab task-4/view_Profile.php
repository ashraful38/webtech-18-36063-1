<!DOCTYPE html>
<html>
<head>
	<title>Public home design</title>
</head>
<body>
 <table border="1"  width="100%">
  	<tr>

 		<td colspan="3" align="right">
 		    <a href="public_home.php"> <img src="download.png" align="left"> </a>
			<?php 


session_start();

if (isset($_SESSION['un'])) {
	//echo "<p1> Welcome ".$_SESSION['un']."</p1>";
	echo "<a href='Dashboard.php'>Login in as ".$_SESSION['un']."</a>" ;
	echo "<br><a href='logout.php'>Logout</a><br>";

}


 ?>
		

 		</td>
 	</tr>
 	<tr height="300px">
 		<td colspan="2" style="width: 30%" >
				<fieldset>
 				<table border="0">
 					<td height="300px" width="450px">
 						<b>Account</b>
 						<hr>
							<ul><a href="Dashboard.php"> Dashbord </a> </ul>
							<ul><a href="view_Profile.php"> View Profile </a> </ul>
							<ul><a href="edit_profile.php"> Edit Profile </a> </ul>
							<ul><a href="Profile_picture.php"> Change Profile Picture </a> </ul>
							<ul><a href="change_Password.php"> Change Password </a> </ul>
							<ul><a href="Logout.php"> Logout </a> </ul>
 				
 				</table>
 			</fieldset>
		</td>
		<td>
			<form >
 				<fieldset style="  width: 50% ;max-width: 50%;height:100%">
 					<legend>PROFILE</legend>
 					  <div class="container" style="width:500px;">              
                <div class="table-responsive"> 
 						<table  class="table table-bordered">
 						<tr>  
                              <th>Name</th>
                               <th>E-mail</th>  
                               <th>User name</th>
                               <th>Gender</th><br>   
                               <th>Date of birth</th>  
                                
                                 
                          </tr>  
                          <?php   
                          $data = file_get_contents("data.json");  

                          $data = json_decode($data, true);  
                
                          foreach($data as $row)  
                          {  
                          	 echo '<tr>
                              
                               <td>'.$row["name"].'</td>
                               <td>'.$row["e-mail"].'</td>
                               <td>'.$row["username"].'</td>
                               <td>'.$row["gender"].'</td>
                               <td>'.$row["dob"].'</td>
                               </tr>';  
                          }  
	                          
                            
                          ?>  
                          <td><img src="user.png" height="150" width="150"><br></td>
	 							<td><a href=""> Change </a></td>
 							</tr>
 							<tr>
	 							<td colspan="2"><a href="edit_profile.html"> Edit Profile </a></td>
	 							
 							</tr>
 							
 								
 						</table>
 						 </div>
                 </div>
 								
 				</fieldset>
 			</form>

		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>