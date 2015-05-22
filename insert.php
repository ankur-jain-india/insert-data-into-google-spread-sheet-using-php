<?php

include 'spreadsheet.php';
$Spreadsheet = new Spreadsheet("demo@gmail.com", "password");
$Spreadsheet->
        setSpreadsheet("information")->
        setWorksheet("Sheet1")->
        add(array("Header 1" => "Cell1", "Header 2" => "Cell2"));
?>
