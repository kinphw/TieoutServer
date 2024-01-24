<?php       
        #exec("excel a.xlsx", $a, $b);    
        $path = './status.txt';

        if(isset($_GET['how'])) {
            if($_GET['how'] == 'direct') {
                $path = '../status.txt';
            };
        };
        
        $tmp = file_exists($path);
        
        if($tmp) {
            $file = fopen($path, 'r');
            $msg = fgets($file);
            #echo $msg;
            fclose($file);
        } else {
            echo "status CALL Failed";
        };

        echo $msg;        
?>