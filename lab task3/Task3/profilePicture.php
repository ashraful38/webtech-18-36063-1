<?php
	
	if(isset($_POST['Submit']))
	{
		
		if(getimagesize($_POST['image']) > 4)
		{
			echo "Choose Less tham 4MB";
		}
		
	}
?>







<!DOCTYPE html>
<html>
<head>
	<title>Profile Picture</title>
</head>
<body>

<form>
	<fieldset>
		<legend>PROFILE PICTURE</legend>
		<table>
			<tr>
				<td>
					Picture <input type="file" name="image" value="">
				</td>
				
			</tr>
		</table>
		<hr>
		<input type="submit" name="submit" value="Submit">
		</hr>
	</fieldset>
</form>
</body>
</html>