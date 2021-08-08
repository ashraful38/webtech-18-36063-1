<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Online Tutor Finder</title>
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">

  <link rel="stylesheet" href="SStyle.css">

</head>

<body>


  <nav class="navbar navbar-expand-sm navbar-dark bg-dark ">
    <a href="#" class="navbar-brand"> ONLINE TUTOR FINDER </a>
    <ul class="navbar-nav ml-auto">
      <li class="nav-item"><a href="1home.php" class="nav-link "> Home </a></li>
      <li class="nav-item"><a href="login.php" class="nav-link "> Login </a></li>
      <li class="nav-item"><a href="registration.php" class="nav-link "> Registration </a></li>
    </ul>
  </nav>


  <section class="form_holder d-flex justify-content-center align-items-center">
    <form class="login_form p-5" action="loginCheck.php" onsubmit="return validation()" method="POST">
      <h2>Please Log In</h2>
      <div class="form-group">
        <label for="username">Username</label>
        <input type="text" class="form-control" id="txtMail" type="text" name="username" placeholder="Username">
        <span id="userMail"></span>
      </div>
      <div class="form-group">
        <label for="password">Password</label>
        <input type="password" class="form-control" id="Ipassword" type="password" name="password" placeholder="Password">
        <span id="userPass"></span>
      </div>
      <div class="form-check">
        <input type="checkbox" class="form-check-input" id="exampleCheck1">
        <label class="form-check-label" for="exampleCheck1">Check me out</label></br>
      </div>
      <button type="submit" class="btn btn-success mt-1" name="submit">Submit</button>
      <div>
        <ul id="forgotPassword">
          <li id="fPassword"><a id="fpref" href="#">ForgotPassword? </a> </li>
          <li id="sUp"><a id="suref" href="#">Sing UP</li>

        </ul>
      </div>
    </form>
  </section>





  <script type="text/javascript" src="login.js"> </script>
  <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>

</body>

</html>