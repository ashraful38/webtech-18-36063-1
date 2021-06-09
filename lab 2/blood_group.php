<?php
?>
<html>
	<head>
	<title>Blood Group</title>
	</head>
	
		<body >
			<form method="post" action="blood_Group_process.php">
			<fieldset>
				<legend><b>BLOOD GROUP</b></legend>
					 <select>
						<option value="A+" name="blood_group" checked="unchecked">A+</option>
						<option value="A-" name="blood_group" checked="unchecked">A-</option>
						<option value="B+" name="blood_group" checked="unchecked">B+</option>
						<option value="B-" name="blood_group" checked="unchecked">B-</option>
						<option value="AB+" name="blood_group" checked="unchecked">AB+</option>
						<option value="AB-" name="blood_group" checked="unchecked">AB-</option>
						<option value="O+" name="blood_group" checked="unchecked">O+</option>
						<option value="O-" name="blood_group" checked="unchecked">O-</option>
					</select> 
					<hr></hr>
					<input type="submit" value="Submit" name="submit"/>
			</fieldset>
			</form>
		</body>
		
</html>