<!DOCTYPE html>
<html>
<head>
	 <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
	<title>Logout</title>
</head>
<body>
	<header>
 <?php include 'design.php';?>
 <div class="topnav">
	<?php
session_start();
session_destroy();
echo "Logout Sucessful.<br> click here to <a class='btn btn-outline-success' href='login.php'> Login again </a> ";
?>
 </div>
</header>
 <br>

</body>
</html>
