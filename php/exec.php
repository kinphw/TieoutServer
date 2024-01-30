<?php

if (!(file_exists("c:/projects/tieout/workspace/a.htm"))) {
    echo "No a.htm";
} elseif (!(file_exists("c:/projects/tieout/workspace/b.htm"))) {
    echo "No b.htm";
} else {
    exec("excel c:/projects/tieout/tieout.xlsb");
    echo "Tie-out done. pls download the resulf file";
}
?>