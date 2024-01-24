<?php

if(isset($_COOKIE["ExecFlag"])) {
    echo "1분에 1번만 클릭할 수 있습니다. (과부하 방지)";        
} else {
    $flag = setcookie("ExecFlag", 1, time()+60, "/");
    if($flag) {
        echo "COOKIE_SUCCESS";

    } else {
        echo "쿠키설정안됨";
    }
}
?>