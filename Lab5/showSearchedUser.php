<!DOCTYPE html>
<html>
<head>
	<title>search user </title>
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
       
<table>
	<thead>
		<tr>
			<th>User_id</th>
			<th>User_name</th>
			<th>Action</th>
		</tr>
	</thead>
	<tbody>
		<?php foreach ($allSearchedUsers as $i => $user): ?>
			<tr>
				<td><a href="../showStudent.php?id=<?php echo $user['ID'] ?>"><?php echo $user['ID'] ?></a></td>
				<td><?php echo $user['Username'] ?></td>
				<td><a href="../editStudent.php?id=<?php echo $user['ID'] ?>">Edit</a>&nbsp<a href="deleteStudent.php?id=<?php echo $user['ID'] ?>">Delete</a></td>
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