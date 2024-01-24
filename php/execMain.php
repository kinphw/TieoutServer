<?php

if(isset($_GET["ExecMain"])) {
    exec('excel c:/projects/tieout/tieout.xlsb');
    echo ("배치완료됨");
} else {
    echo "FAIL-execMain";
}

?>