<!DOCTYPE html>
<html>
<head>
  <style >
    .main {
  margin-left: 160px; 
  font-size: 28px; 
  padding: 0px 10px;
}

  </style>
</head>
<body>
 
   <?php include 'design.php';?>
  <header>
  <h2>Online Tutor Finder</h2>
  <h1 ></h1>  
  
  <div class="topnav">
    
     <?php
      session_start();


echo "Logged in as ". $_SESSION['email'];
echo " | <a  class='btn btn-outline-danger' href='logout.php'> Logout </a>" ;
?>
 
</div>
</header>
  <?php include 'sidebar.php';?>

  
</div>
 <div class="main">
    <h1>Welcome <?php 
   
     echo " ". $_SESSION['name'];

    ?></h1>
    <br><br><br>
    
  </div>
<footer>
  <p>Copyright <span>&#169;</span>Ashraful Alam</p>
</footer>
</body>
</html>