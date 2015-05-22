<?php

include 'spreadsheet.php';
$Spreadsheet = new Spreadsheet("demo@gmail.com", "password");

$contents = $Spreadsheet->
        setSpreadsheet("information")->
        setWorksheet("Sheet1")->
        read();
// All of the examples below produce the same output assuming column 1 row 1 is named "Header 1"

echo $contents[0][0] . "<br>";  // flat numeric access
echo $contents["A2"] . "<br>";  // associative table access
echo $contents["Header 1"][1];  // associative named access
?>
