<?php       
        #exec("excel a.xlsx", $a, $b);    
        $path = $_SERVER['DOCUMENT_ROOT'].'/tieout/status.txt';
        
        if(file_exists($path)) {

            $file = fopen($path, 'r');
            $msg = fgets($file);
            $msg = iconv("cp949","utf-8", $msg); #240313
            
            fclose($file);
        } else {
            echo "No status file...";
        };

        echo $msg;        
?>