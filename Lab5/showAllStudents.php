<!DOCTYPE html>
<html>
<head>
	<title>show student </title>
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

$students = fetchAllStudents();


    include "nav.php";



?>
 				
 				</table>
 			</fieldset>
		</td>
		<td align="center">
		<table>
	<thead>
		<tr>
			<th>Name</th>
			<th>Surname</th>
			<th>Username</th>
			<th>Password</th>
			<th>Image</th>
			<th>Action</th>
		</tr>
	</thead>
	<tbody>
		<?php foreach ($students as $i => $student): ?>
			<tr>
				<td><a href="showStudent.php?id=<?php echo $student['ID'] ?>"><?php echo $student['Name'] ?></a></td>
				<td><?php echo $student['Surname'] ?></td>
				<td><?php echo $student['Username'] ?></td>
				<td><?php echo $student['Password'] ?></td>
				<td><img width="100px" src="uploads/<?php echo $student['image'] ?>" alt="<?php echo $student['Name'] ?>"></td>
				<td><a href="editStudent.php?id=<?php echo $student['ID'] ?>">Edit</a>&nbsp<a href="controller/deleteStudent.php?id=<?php echo $student['ID'] ?>" onclick="return confirm('Are you sure want to delete this ?')">Delete</a></td>
			</tr>
		<?php endforeach; ?>
		

	</tbody>
	

</table>

		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>