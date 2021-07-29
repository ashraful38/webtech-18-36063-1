<!DOCTYPE html>
<html>
<head>
  <title></title>
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
echo " | <a class='btn btn-outline-danger' href='logout.php'> Logout </a>" ;

?>
 
</div>
</header>
  <?php include 'sidebar.php';?>

  
</div>
<div id="main-content" class="main">
    <h2>Update Record</h2>
     <?php include 'dbcon.php';?>
    <?php 
     $user_id=$_GET['id'];
    $sql="SELECT * FROM users WHERE id={$user_id}";
    $result=mysqli_query($con, $sql) or die("query failed");

  

  if (mysqli_num_rows($result)>0) {
      while ($row=mysqli_fetch_assoc($result)) {
    
    ?>
    <form class="post-form" action="editedData.php" method="post">
       <table>
    
      <tr >

        <td >Name</td>
        <input type="hidden" name="id" value="<?php echo $row['id'];?>"/>

        <td  > <input type="text" name="name" value="<?php echo $row['name'];?>"/></td>
       
         
      </tr>

        <tr>
        <td > User Name   </td>
       

        <td><input type="text" name="username" value="<?php echo $row['username'];?>"/></td>
          
      </tr>   

      <tr>
        <td>Email</td>

        <td><input type="text" name="email" value="<?php echo $row['email'];?>"/></td>
        
      </tr>
     

    </table>
      <input class="btn btn-info" type="submit" value="Update"/>
    </form>
    <?php } 
      }
    ?>
</div>

<br>
<footer>
  <p>Copyright <span>&#169;</span>Ashraful Alam</p>
</footer>
</body>
</html>