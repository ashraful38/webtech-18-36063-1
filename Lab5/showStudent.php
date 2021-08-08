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
require_once 'controller/studentInfo.php';

$student = fetchStudent($_GET['id']);


    include "nav.php";



?>

 				
 				</table>
 			</fieldset>
		</td>
		<td align="center">
		<table>
	<tr>
		<th>Name</th>
		<th>Surname</th>
		<th>Username</th>
		<th>Password</th>
		<th>Image</th>
	</tr>
	<tr>
		<td><a href="showStudent.php?id=<?php echo $student['ID'] ?>"><?php echo $student['Name'] ?></a></td>
		<td><?php echo $student['Surname'] ?></td>
		<td><?php echo $student['Username'] ?></td>
		<td><?php echo $student['Password'] ?></td>
		<td><img width="100px" src="uploads/<?php echo $student['image'] ?>" alt="<?php echo $student['Name'] ?>"></td>
	</tr>

</table>


		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>