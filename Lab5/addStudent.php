<!DOCTYPE html>
<html>
<head>
	<title>add </title>
</head>
<body>
 <table border="1"  width="100%">
 	<tr>

 	
     <td colspan="3" align="center">
 		    <a href="public_home.php"> <h1 >Online Tutor Finder</h1> </a>
		

 		</td>
 	</tr>
 	<tr height="300px">
 		<td colspan="2" style="width: 30%" >
				<fieldset>
 				<table border="0">
 					<td height="300px" width="450px">
 						<b>Account</b>
 						<hr>
                         <?php 
        include "nav.php";

     ?>

 				
 				</table>
 			</fieldset>
		</td>
		<td align="center">
        <form action="controller/createStudent.php" method="POST" enctype="multipart/form-data">
  <label for="name">Name:</label><br>
  <input type="text" id="name" name="name"><br>
  <label for="surname">Surname:</label><br>
  <input type="text" id="surname" name="surname"><br>
  <label for="username">User Name:</label><br>
  <input type="text" id="username" name="username"><br>
  <label for="password">Password:</label><br>
  <input type="password" id="password" name="password"><br>
  <input type="file" name="image"><br><br>
  <input type="submit" name = "createStudent" value="Create">
  <input type="reset"> 
</form>

		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>