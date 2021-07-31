<?php

include 'conn.php';

if(isset($_POST['submit']))
{
  $userNamne=$_POST['username'];
  $password=$_POST['password'];

  $q = "SELECT * FROM `registration` WHERE email='$userNamne' and password='$password'";

  $query = mysqli_query($con, $q);

  if($query->num_rows > 0){
    header("location: NewsFeed.php");
    echo "hi";
  }
  else{
      echo "Invalid email or password";
      header("location: login.php");
  }

}


?>