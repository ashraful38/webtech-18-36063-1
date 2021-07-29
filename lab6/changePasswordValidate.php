<?php

if (isset($_POST['submit'])) {

  if(  empty($_POST['new_password']) || empty($_POST['con_password'])){
            echo"Plaese fill up all the password field";

        }
 elseif(($_POST['new_password']) == ($_POST['con_password'])) {
 	
	$user_id=$_POST['id'];
 $user=$_POST['name'];
 $currPassword=$_POST['password'];
 $newPassword= $_POST['new_password'];
 $confrimPassword= $_POST['con_password'];
   $con=mysqli_connect("localhost","root","","user") or die("failed");
    $sql="UPDATE `users` SET `password`='$newPassword',`curr_password`='$confrimPassword' WHERE `id` ={$user_id}";     

   $result=mysqli_query($con, $sql) or die("query failed");
   header('Location: profile.php');
  mysqli_close($con);
}
else{
	echo "new password and confrim password not matched !";

}

}


?>


