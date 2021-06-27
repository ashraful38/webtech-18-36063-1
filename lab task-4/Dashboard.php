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
							<ul><a href="Gedit_profile.html"> Edit Profile </a> </ul>
							<ul><a href="HprofilePicture.html"> Change Profile Picture </a> </ul>
							<ul><a href="Ichange_Password.html"> Change Password </a> </ul>
							<ul><a href="login.html"> Logout </a> </ul>
 				
 				</table>
 			</fieldset>
		</td>
		<td align="center">
			 <?php 

             if (isset($_SESSION['un'])) {
	           echo "<h1> Welcome ".$_SESSION['un']."</h1>";

             }
            ?>

		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>