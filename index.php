<?php

/**
 * A quick guide for Jafar Ahmed 
 *
 * This script helps to understand how to read xlsx file using PHPExcel library 
 * specifically reading column wise and row wise 
 *
 * PHP version 5
 *
 *
 * @category   php exel report
 * @author     Original Author <hi@ponick.me>
 *
 */




error_reporting(E_ALL);
set_time_limit(0);

date_default_timezone_set('Asia/Dhaka');

?>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

<title>XL Report</title>

</head>
<body>


<?php

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . 'Classes/');

/** PHPExcel_IOFactory */
include 'PHPExcel/IOFactory.php';


$inputFileName = './sampleData/m.xlsx';

$objPHPExcel = PHPExcel_IOFactory::load($inputFileName);



$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
$highestRow = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
$highestColumm = $objPHPExcel->setActiveSheetIndex(0)->getHighestColumn();

$columnwiseEx1 = $objPHPExcel->getActiveSheet()->rangeToArray('A1:AC1');
$columnwiseEx2 = $objPHPExcel->getActiveSheet()->rangeToArray('A3:AC3');


//example of getting specific cell value like date
$date_of_report=$objPHPExcel->getActiveSheet()->rangeToArray('E1:F1');
for($i=0; $i<1;$i++){
	if($date_of_report[0][$i]!=""){
$rDate=$date_of_report[0][$i];	
}
}

echo"<h2>Date: $rDate</h2>";

/*

//column wise reading 
for($i=0; $i<20;$i++){
	if($columnwiseEx1[0][$i]!=""){
echo $columnwiseEx1[0][$i]." | ";	
}
}
echo"<br/>";
for($i=0; $i<20;$i++){
	if($columnwiseEx2[0][$i]!=""){
echo $columnwiseEx2[0][$i]." | ";	
}
}

*/
echo"<h2>Territory Name</h2>";

///territory name
$ttn = $objPHPExcel->getActiveSheet()->rangeToArray('B5:B2000');
for($i=0; $i<$highestRow;$i++){
	if($ttn[$i][0]!=""){
echo $ttn[$i][0]." <br/> ";	
}
}

echo"<h2>Territory Code</h2>";

///territory code
$ttc = $objPHPExcel->getActiveSheet()->rangeToArray('C5:C2000');
for($i=0; $i<$highestRow;$i++){
	if($ttc[$i][0]!=""){
echo $ttc[$i][0]." <br/> ";	
}
}


?>
<body>
</html>