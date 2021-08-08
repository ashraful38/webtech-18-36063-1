
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
          <!-- [SEARCH FORM] -->
    <form method="post" action="controller/findUser.php">
      <h1>SEARCH FOR USERS</h1>
      <input type="text" name="user_name" />
      <input type="submit" name="findUser" value="Search"/>
    </form>

		</td>
 		
 	</tr>

 	<tr>
 		<td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
 		
 	</tr>
 </table>
 
</body>
</html>