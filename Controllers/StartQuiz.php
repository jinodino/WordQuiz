<?php

require_once "../Dao/ExcelProcessClass.php";

$ExcelProcessObj = new ExcelProcess();

$examArr = json_encode($ExcelProcessObj->submitAnExamination());

echo $examArr;

?>