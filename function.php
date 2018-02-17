<?php

// fungsi styling cell
function cellStyle($excel, $cell, $style)
{
   $excel->getActiveSheet()->getStyle($cell)->applyFromArray($style);
}

// fungsi cell width
function cellWidth($excel, $column, $width)
{
   $excel->setActiveSheetIndex(0)->getColumnDimension($column)->setWidth($width);
}

// cell background
$title = array( 
   'fill' => array( 
      'type' => PHPExcel_Style_Fill::FILL_SOLID, 
      'color' => array('rgb' => 'f0f0f0')
   ), 
   'font' => array( 
      'bold' => true, 
      'size' => 10, 
      'name' => 'Arial'
   )
);


?>
