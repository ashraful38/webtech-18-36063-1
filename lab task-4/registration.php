<!DOCTYPE html>
<html>
<head>
  <title>Public Home</title>
</head>
<body>
<table border="1" width="100%">
  <tr>
    <td align="right">
       <a href="home.html"> <img src="download.png" align="left"> </a>
       <a href="public_home.php"> Home </a> |
      <a href="login.php"> Login </a> |
      <a href="registration.php"> Registration </a> 
      
    </td>
  </tr>
    <tr height="200px">
    <td colspan="2" align="center">
    <?php  
 $message = '';  
 $error = '';  
 if(isset($_POST["submit"]))  
 {  
      if(empty($_POST["name"]))  
      {  
           $error = "<h1 class='text-danger'>Enter Name</h1>";  
      }
  else {
    $name = $_POST["name"];
   
     if (!preg_match("/^[a-zA-Z-'- ]*$/",$name)) {
      $error = "Only letters and white space allowed";
      $name = "";
    }
  }




      if(empty($_POST["email"]))  
      {  
           $error = "<label class='text-danger'>Enter an e-mail</label>";  
      }
       else {
    $email = test_input($_POST["email"]);
    
     if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
      $error = "Invalid email format";
      $email = "";
    }
  }





       if(empty($_POST["un"]))  
      {  
           $error = "<label class='text-danger'>Enter a username</label>";  
      }  
      if(empty($_POST["pass"]))  
      {  
           $error = "<label class='text-danger'>Enter a password</label>";  
      }

       if (test_input($_POST["pass"])<8) 
          {
             $error = "<Password must not be less than eight (8) characters>";
          }




       if(empty($_POST["Cpass"]))  
      {  
           $error = "<label class='text-danger'>Confirm password field cannot be empty</label>";  
      } 
       if(empty($_POST["gender"]))  
      {  
           $error = "<label class='text-danger'>Gender cannot be empty</label>";  
      } 
       
       
      else  
      {  
           if(file_exists('data.json'))  
           {  
                $current_data = file_get_contents('data.json');  
                $array_data = json_decode($current_data, true);  
                $extra = array(  
                     'name'               =>     $_POST["name" ],  
                     'e-mail'          =>     $_POST["email"],  
                     'username'     =>     $_POST["un"],
                     //'password'   =>      $_POST["pass"],
                     'gender'    =>     $_POST["gender"],  
                     'dob'     =>     $_POST["dob"]  
                );  
                $array_data[] = $extra;  
                $final_data = json_encode($array_data);  
                if(file_put_contents('data.json', $final_data))  
                {  
                     $message = "<label class='text-success'>File Appended Success fully</p>";  
                }  
           }  
           else  
           {  
                $error = 'JSON File not exits';  
           }  
      }  
 } 


 function test_input($data) {
  $data = trim($data);
  $data = stripslashes($data);
  $data = htmlspecialchars($data);
  return $data;
} 
 ?>  
 <!DOCTYPE html>  
 <html>  
      <head>  
           
           
      </head>  
      <body>  
           <br />  
           <div class="container" style="width:500px;">  
                                
                <form method="post">  
                     <?php   
                     if(isset($error))  
                     {  
                          echo $error;  
                     }  
                     ?>  
                     <br /> 
                      <fieldset> 
                     <legend>REGISTRATION</legend> 
                     <label>Name</label>  
                     <input type="text" name="name" class="form-control" /><br />  
                     <label>E-mail</label>
                     <input type="text" name = "email" class="form-control" /><br />
                     <label>User Name</label>
                     <input type="text" name = "un" class="form-control" /><br />
                     <label>Password</label>
                     <input type="password" name = "pass" class="form-control" /><br />
                     <label>Confirm Password</label>
                     <input type="password" name = "Cpass" class="form-control" /><br />

                    <fieldset>
                    <legend>Gender</legend>
                    <input type="radio" id="male" name="gender" value="male">
                     <label for="male">Male</label>                     
                     <input type="radio" id="female" name="gender" value="female">
                     <label for="female">Female</label>
                     <input type="radio" id="other" name="gender" value="other">
                     <label for="other">Other</label><br>
                     </fieldset>



                     <fieldset> 
                     <legend>Date of Birth:</legend>
                     <input type="date" name="dob"> <br><br>
                    </fieldset> 
                    </fieldset>
                     
                     <input type="submit" name="submit" value="Submit" class="btn btn-info" /><br />                      
                     <?php  
                     if(isset($message))  
                     {  
                          echo $message;  
                     }  
                     ?>  
                </form>  
           </div>  
           <br />  
      </body>  
 </html>  
    </td>
    
  </tr>

  <tr>
    <td colspan="3" align="center" height ="50px">Copyright @ 2021</td>
    
  </tr>
 </table>
</body>
</html>