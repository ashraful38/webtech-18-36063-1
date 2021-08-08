<!DOCTYPE html>
<html>
<head>
	<title>edit student </title>
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
             <?php 

require_once 'controller/studentInfo.php';
$student = fetchStudent($_GET['id']);


 include "nav.php";



 ?>

 				
 				</table>
 			</fieldset>
		</td>
		<td align="center">
    <form action="controller/updateStudent.php" method="POST" enctype="multipart/form-data">
  <label for="name">Name:</label><br>
  <input value="<?php echo $student['Name'] ?>" type="text" id="name" name="name"><br>
  <label for="surname">Surname:</label><br>
  <input value="<?php echo $student['Surname'] ?>" type="text" id="surname" name="surname"><br>
  <label for="username">User Name:</label><br>
  <input value="<?php echo $student['Username'] ?>" type="text" id="username" name="username"><br>
  <input type="file" name="image"><br><br>
  <input type="hidden" name="id" value="<?php echo $_GET['id'] ?>">
  <input type="submit" name = "updateStudent" value="Update">
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