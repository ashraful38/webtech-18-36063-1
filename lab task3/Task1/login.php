


<?php
	session_start();

	if(isset($_POST['submit']))
	{
		$username = $_POST['username'];
		$password = $_POST['password'];

		if(strlen($username)<2)
		{
			echo "Please Enter more than 2 charactres";
		}
		elseif (strlen($password)<8) 
		{
			echo "Password must not be less than eight (8) characters";
		}
		/*else
		{
		

			if($password == '@' || $password == '#'|| $password == '$'|| $password == '%')
			{
				
				echo "Thank You";

			}
			else
			{
				echo "Password must contain at least one of the special characters (@, #, $, %)";
			}
		}*/

		
	}



?>


<!DOCTYPE html>
<html>
<head>
	<title>Login Page</title>
</head>
<body>

	<form method="post" action="loginCheck.php"> 
		<fieldset>
			<legend>LOGIN</legend>
			<table>
				<tr>
					<td>User Name :</td>
					<td><input type="text" name="username"></td>
				</tr>
				<tr>
					<td>Password :</td>
					<td><input type="password" name="password"></td>

				</tr>

			</table>

				<hr>
				<input type="checkbox" name="checkbox" value="Remember Me">Remember Me <br>
				<input type="submit" name="submit" value="Submit">
					
				<a href="forgotPassword.html"> Forget Password?
				
			
		</fieldset>
	</form>
</body>
</html>