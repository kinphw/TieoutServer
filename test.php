<?php
    setcookie("ExecFlag", 1, time()+60, "/");

    if(isset($_COOKIE["ExecFlag"])) {
        echo "세팅됨";        
    } else {
        echo "세팅안됨";
    }
?>