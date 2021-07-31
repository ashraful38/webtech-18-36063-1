<?php

$con = mysqli_connect('localhost', 'root');

mysqli_select_db($con, 'nproject');

if($con)
{
    //echo "database connected";
}
else{
    echo "dataabase not connected";
}



?>