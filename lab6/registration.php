 
<!DOCTYPE html>  
 <html>  
      <head>  
            
           <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>  
           
           <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>  


          
      </head>  
      <body>  
<?php include 'design.php';?>
<?php  
include 'dbcon.php';
 if (isset($_POST['submit'])) {

if(empty($_POST['name']) || empty($_POST['username']) || empty($_POST['email']) || empty($_POST['password']) || empty($_POST['curr_password'])){
            echo"Plaese fill out all the field";
        }

else{
  $name=mysqli_real_escape_string($con,$_POST['name']);
  $username=mysqli_real_escape_string($con,$_POST['username']);
  $email=mysqli_real_escape_string($con,$_POST['email']);
  $password=mysqli_real_escape_string($con,$_POST['password']);
   $curr_password=mysqli_real_escape_string($con,$_POST['curr_password']);
   $emailquery= "SELECT * FROM users WHERE email='$email'";
   $query=mysqli_query($con,$emailquery);

   $emailcount=mysqli_num_rows($query);

   if ($emailcount>0) {
     echo "email already exist! ";
   }else{
    if ($password === $curr_password) {
      $insertquery="INSERT INTO users( name, username, email,password, curr_password) VALUES('$name','$username','$email','$password','$curr_password')";
      $iquery=mysqli_query($con,$insertquery);
      echo "registration suessful";
      header('location:login.php');

    }else{
      echo "password not matching";
     
   }  
   }
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
           <br />  
           <div >  
                              
                <form method="post">  
                      
                     <br />  
                      <fieldset>
    <legend> <b> <font size="5px">Customer Registration  </font> </b> </legend>
    <fieldset>

    <table>
    
      <tr >

        <td>Name</td>

        <td  ><input type="text" name="name"/></td>
       
         
      </tr>

        <tr>
        <td > User Name   </td>
       

        <td><input type="text" name="username"/></td>
          
      </tr>   

      <tr>
        <td>Email</td>

        <td><input type="text" name="email"></td>
        
      </tr>
      

      <tr>
        <td>Password </td>
        
        <td><input type="password" name="password"/></td> <br>
        
      </tr>

      <tr>
        <td>Confrim Password </td>
       
        <td><input type="password" name="curr_password"/></td> <br>
      </tr>



    </table>
    </fieldset>
        
    <hr>
    
    <br>
    <table>
      <tr>
        <td>
          <input class="btn btn-outline-secondary" type="submit" name="submit" value="Submit"> 
     <input class="btn btn-outline-dark" type="reset" value="Reset">
        </td>
      </tr>

    </table>
    
  </fieldset>                   
                      
                  
                </form>  
           </div>  
           <br /> 
            <footer>
  <p>Copyright <span>&#169;</span> Ashraful Alam</p>
</footer> 
      </body>  
 </html>  