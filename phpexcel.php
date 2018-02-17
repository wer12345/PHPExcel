<?php
require_once './Classes/PHPExcel.php';
require_once './function.php';

// Create PHPExcel object
$excel = new PHPExcel();

// remove gridlines
$excel->getActiveSheet()->setShowGridlines(false);

// Excel Template
excelTemplate($excel);

// input cell
data($excel);

// Redirect to browser Download
//
header('Content-Tyype: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="text.xlsx"');
header('Cache-Control: max-age=0');

// Write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
// Output to php output
$file->save('php://output');

?>
