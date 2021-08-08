<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Search</title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">

<link rel="stylesheet" href="SStyle.css">

</head>
<body>

<nav class="navbar navbar-expand-sm navbar-dark bg-dark ">
<a href="#" class="navbar-brand"> ONLINE TUTOR FINDER </a>
<ul class="navbar-nav ml-auto">
    <li class="nav-item"><a href="#" class="nav-link "> NEWS FEED </a></li>
    <li class="nav-item"><a href="Search.php" class="nav-link "> SEARCH </a></li>
    <li class="nav-item"><a href="post.php" class="nav-link "> POST </a></li>
    <li class="nav-item"><a href="#" class="nav-link "> NOTIFICATION </a></li>
    <li class="nav-item"><a href="#" class="nav-link "> PROFILE </a></li>
    <li class="nav-item"><a href="#" class="nav-link "> LOGOUT </a></li>



 
</ul>


</nav>


<section class="Search_holder1 d-flex justify-content-center align-items-center">
<form class="Search_form p-5" action="#" method="POST">
	     
                    <h1 style="margin-bottom: 10px"><label for="">Search Tutor!</label></h1>
					</br>
               
				
				<div class="Subject">
					<label for="">Subject :</label>
					<select name="subject">
					<label for="">  <option  value="Bangla"> Bangla </option></label>
					<label for="">  <option value="English"> English </option></label>
					<label for="">  <option value="Math"> Math </option></label>
					 <label for=""> <option value="Physic"> Physic </option></label>
					 <label for=""> <option value="ICT"> ICT </option></label>
					 <label for=""> <option value="Chemistry"> Chemistry </option></label>
					 <label for=""> <option value="Biology"> Biology </option></label>
					 <label for=""> <option value="Religion"> Religion </option></label>
					 <label for=""> <option value="General Science"> General Science </option></label>
					 <label for=""> <option value="Economics"> Economics </option></label>
					 <label for=""> <option value="Statistics"> Statistics </option></label>
					 <label for=""> <option value="Social Science"> Social Science </option></label>
					<label for="">  <option value="Finance"> Finance </option></label>
					 <label for=""> <option value="History"> History </option></label>
					  <label for=""><option value="Accounting"> Accounting </option></label>
					  </select>
                    
					  
						<label for="">Class:</label>
						<select name="class">
						  <label for=""><option value="One">One</option></label>
						  <label for=""><option value="Two">Two</option></label>
						  <label for=""><option value="Three">Three</option></label>
						  <label for=""><option value="Four">Four</option></label> 
						  <label for=""><option value="Five">Five</option></label> 
						  <label for=""><option value="six">six</option></label> 
						  <label for=""><option value="Seven">Seven</option></label> 
						  <label for=""><option value="Eight">Eight</option> </label>
						 <label for=""> <option value="Nine">Nine</option> </label>
						  <label for=""><option value="Ten">Ten</option> </label>
						 <label for=""> <option value="Eleven">Eleven</option></label>
						 <label for=""> <option value="Twelve">Twelve</option></label>  
						    </select>
						
						<label for="">Medium :</label>
						<select name="medium">
					 <label for=""> <option value="Bangla"> Bangla </option></label>
					 <label for=""> <option value="English"> English </option></label>
					 <label for=""> <option value="Others"> Others </option></label>
					 <span id="checkMedium"></span>
					  </select>
					  </div>
					
					<br>
					
				<div class="Salary_range">
                <th align="left">
				<label for="">Salary Range : </label> 
				<select name="salRangeLow">
					          <label for=""><option value="1000">1000</option></label>
							  <label for=""><option value="2000">2000</option></label>
							 <label for=""> <option value="3000">3000</option></label>
							 
				</select>
				<?php echo"&nbsp"."<label> To </label>"."&nbsp";?>
				<select name="salRangeHigh">
					         <label for=""> <option value="2000">2000</option></label>
							 <label for=""> <option value="3000">3000</option></label>
							 <label for=""> <option value="4000">4000</option></label>
							 <label for=""> <option value="5000">5000</option></label>
							 <label for=""> <option value="6000">6000</option></label>
							 <label for=""> <option value="7000">7000</option></label>
							 
				</select>
           </div>
 			
             
					<br>
					<div class="Location">
			   <label for="">Location:</label>
						  <select name="location">
						 <label for=""> <option value="Badda">Badda</option></label>
						 <label for=""> <option value="Nikunjo">Nikunjo</option></label>
						 <label for=""> <option value="Basabo">Basabo</option></label>
						 <label for=""> <option value="Gulshan">Gulshan</option></label>
						 <label for=""> <option value="Mohammodpur">Mohammodpur</option></label>
						  <label for=""><option value="Banani">Banani</option></label>
						<label for=""> <option value="khilkhet">khilkhet</option></label>	
							
						  </select>
                               </div>
						
					<br>
				<div class="Preferred_Institution">
				<label for="">Preferred Institution:</label>
				<select name="institution">
							<label for=""><option value="Buet">Buet</option></label>
					         <label for=""> <option value="Aiub">Aiub</option></label>
							  <label for=""><option value="NSU">NSU</option></label>
							 <label for=""> <option value="IUB">IUB</option></label>
							  <label for=""><option value="DU">DU</option></label>
							  <label for=""><option value="Any">Any</option></label>
							  
							 
				</select>
				
					<br><br>
					
				
                <div class="submit_button"><button type="submit" name="submit" >Search</button></div>
           
		</form>
        </section>




<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
</body>
</html>