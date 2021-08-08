<?php

include 'conn.php';

if (isset($_POST['done'])) {
  $userNamne = $_POST['username'];
  $password = $_POST['password'];
  $email = $_POST['email'];
  $phoneNumber = $_POST['phoneNumber'];
  $gender = $_POST['gender'];
  $userType = $_POST['userType'];
  if ($_POST['userType'] == 'student') {
    $status = $pending = 0;
  } elseif ($_POST['userType'] == 'teacher') {
    $status = $pending = 1;
  } elseif ($_POST['userType'] == 'parent') {
    $status = 2;
    $pending = 0;
  }


  $q = "INSERT INTO `registration`(`userName`, `password`, `email`, `phoneNumbe`, `gender`, `userType`, `status`, `pending`) VALUES ('$userNamne','$password','$email','$phoneNumber','$gender','$userType', '$status', '$pending')";

  $query = mysqli_query($con, $q);
}

?>

<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">



  <link rel="stylesheet" href="SStyle.css">
</head>

<body>



  <div>

    <section class="registration_holder1 d-flex justify-content-center align-items-center">
      <form class="registration_form p-4" action="" onsubmit="return validation()" method="POST">
        <h2>Please Registration</h2>
        <div class="form-group ">
          <label for="">Username</label>
          <input type="text" class="form-control" id="RIusername" type="text" name="username" placeholder="Username">
          <span id="fullName"></span>
        </div>
        <div class="form-group ">
          <label for="password">Password</label>
          <input type="password" class="form-control" id="RIpassword" type="password" name="password" placeholder="Password">
          <span id="userPass"></span>
        </div>

        <div class="form-group">
          <label for="confirmpassword"> Confirm Password</label>
          <input type="password" class="form-control" id="CRIpassword" type="password" name="confirmPassword" placeholder="Confirm Password">
          <span id="conPass"></span>

        </div>

        <div class="form-group">
          <label for="email"> Email</label>
          <input type="text" class="form-control" id="ERIpassword" type="text" name="email" placeholder="Enter Email">
          <span id="userMail"></span>
        </div>

        <div class="form-group">
          <label for="phonenumber"> Phone Number</label>
          <input type="text" class="form-control" id="PRIpassword" type="text" name="phoneNumber" placeholder="Enter phonenumber">
          <span id="userPhone"></span>
        </div>

        <div>

          <label id="RLRadio" for="">Gender &nbsp;</label>

          <input type="radio" name="gender" value="male">
          <label id="RLRadio1" for="">Male</label>

          <input type="radio" name="gender" value="female">
          <label id="RLRadio2" for="">Female</label>

          <input type="radio" name="gender" value="other">
          <label id="RLRadio3" for="">Other</label> <br>
          <span id="checkGender"></span>





          <label id="RLRadioType" for="">Register as &nbsp;</label>

          <input type="radio" name="userType" value="student">
          <label id="RLRadioType1" for="">Student</label>

          <input type="radio" name="userType" value="teacher">
          <label id="RLRadioType2" for="">Teacher</label>

          <input type="radio" name="userType" value="parent">
          <label id="RLRadioType3" for="">Parent</label> <br>
          <span id="checkUserType"></span>



        </div>


        <button type="submit" class="btn btn-success mt-1" id="Rsubmit" name="done">Submit</button>
      </form>

    </section>



    <script type="text/javascript" src="registration.js"> </script>
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
    
</body>

</html>