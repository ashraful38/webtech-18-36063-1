<?php
session_start();
?>


<!DOCTYPE html> 
<html> 
<head>
    
</head>
<body>
<?php include 'design.php';?>


<?php
include 'dbcon.php';
if (isset($_POST['submit'])) {
  $email=$_POST['email'];
  $password=$_POST['password'];


  $email_sch="SELECT * FROM users WHERE email='$email' AND password='$password' ";

  

  $query=mysqli_query($con,$email_sch);
  $email_count=mysqli_num_rows($query);

  if ($email_count) {
   $email_pass=mysqli_fetch_assoc($query);
   $db_pass=$email_pass['password']; 
   $_SESSION['email']=$email_pass['email'];
    $_SESSION['name']=$email_pass['name'];

   echo "login sucessful";
   header('location:dashBoard.php');
  }else {
    echo "login failed ";
  }

}
?>
  <header>
  <h2>Online Tutor Finder</h2>
  <h1 ></h1>  
  
  <div class="topnav btn-group" role="group" aria-label="Basic example">
  <a class="btn btn-info" href="Home.html" > Home </a> <br>
  <a class="btn btn-secondary" href="login.php">Login </a><br>
  <a class="btn btn-light" href="registration.php">Registration </a>
 
</div>
</header>
<div >
 
<form method="post" action="<?php echo htmlentities($_SERVER['PHP_SELF']);?>">
    <fieldset>
        <legend> <b> <font size="5px">Login  </font> </b> </legend>
        <table cellpadding="7px" >
            <tr>
                <td > Email   </td>
                

                <td><input type="text" name="email"/></td>
            </tr>

            <tr>
                <td>Password </td>
        
                <td><input type="password" name="password"/></td> <br>
                
            </tr>

        </table>
        <hr>
        
       
        <input class="btn btn-outline-secondary" type="submit" name="submit" value="submit"> 
    
</form>
</div>

<br>
<footer>
  <p>Copyright <span>&#169;</span> Ashraful Alam</p>
</footer>

</body>
</html>