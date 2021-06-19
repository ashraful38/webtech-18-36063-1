<?php
	
	if(isset($_POST['submit']))
	{
		$currentPass = $_POST['currentPassword'];
		$newPass = $_POST['newPassword'];
		$retypeNewPass = $_POST['retypeNPassword'];

		if($currentPass == "" || $newPass =="" || $retypeNewPass =="")
		{
			echo "Null submission";
		}

		elseif( $currentPass  !== $newPass && $newPass == $retypeNewPass)

		{
			echo "Your Password is changed";
		}
		else
		{
			echo "Please put currect values";
		}
	}
?>






<!DOCTYPE html>
<html>
<head>
	<title>Password page</title>
</head>
<body>

	<form > 
		<fieldset>
			<legend>CHANGE PASSWORD</legend>
			<table>
				<tr>
					<td>Current Password :</td>
					<td><input type="password" name="currentPassword"></td>
				</tr>
				<tr>
					<td>New Password :</td>
					<td><input type="password" name="newPassword"></td>

				</tr>
				<tr>
					<td>Retype New Password :</td>
					<td><input type="password" name="retypeNPassword"></td>

				</tr>
			</table>

			<hr>
			<input type="submit" name="submit" value="submit">
		</fieldset>
	</form>
</body>
</html>