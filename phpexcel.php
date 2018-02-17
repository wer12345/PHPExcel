<?php
require_once './Classes/PHPExcel.php';

// Create PHPExcel object
$excel = new PHPExcel();

$excel->getActiveSheet()->getDefaultStyle()->applyFromArray(
   array(
      'borders' => array(
         'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_NONE
         )
      )
   )
);
// Set Column Width
$excel->setActiveSheetIndex(0)->getColumnDimension('A')->setWidth(2);
$excel->setActiveSheetIndex(0)->getColumnDimension('B')->setWidth(35);
$excel->setActiveSheetIndex(0)->getColumnDimension('C')->setWidth(35);

// cell alignment
$excel->getActiveSheet()->getStyle('B2:C45')->applyFromArray(
   array(
      'alignment' => array(
         'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
      )
   )
);

// Merging 
$excel->getActiveSheet()->mergeCells('B6:C6');
$excel->getActiveSheet()->mergeCells('B12:C12');
$excel->getActiveSheet()->mergeCells('B17:C17');
$excel->getActiveSheet()->mergeCells('B22:C22');
$excel->getActiveSheet()->mergeCells('B26:C26');
$excel->getActiveSheet()->mergeCells('B31:C31');
$excel->getActiveSheet()->mergeCells('B37:C37');
$excel->getActiveSheet()->mergeCells('B41:C41');
$excel->getActiveSheet()->mergeCells('E4:F4');

// cell background color
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

$excel->getActiveSheet()->getStyle('B6')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B12')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B17')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B22')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B26')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B31')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B37')->applyFromArray($title);
$excel->getActiveSheet()->getStyle('B41')->applyFromArray($title);


// give border
$excel->getActiveSheet()->getStyle('B5:C45')->applyFromArray(
   array(
      'borders' => array(
         'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'vertical' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         )
      ),
      'font' => array(
         'size' => 10,
         'name' => 'Arial'
      )
   )
);

// Templpate Form Registration
$excel->setActiveSheetIndex(0)
   ->setCellValue('B2', 'Tanggal Register')
   ->setCellValue('B3', 'Tanggal Verif');

$excel->getActiveSheet()->getStyle('B2:C2')->applyFromArray(
   array(
      'borders' => array(
         'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'vertical' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         )
      ),
      'fill' => array(
         'type' => PHPExcel_Style_Fill::FILL_SOLID,
         'color' => array('rgb' => 'ecc0c1')
      ),
      'font' => array(
         'bold' => true,
         'name' => 'Arial',
         'size' => 10
      )
   )
);

$excel->getActiveSheet()->getStyle('B3:C3')->applyFromArray(
   array(
      'borders' => array(
         'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'vertical' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         ),
         'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'a6a6a6')
         )
      ),
      'fill' => array(
         'type' => PHPExcel_Style_Fill::FILL_SOLID,
         'color' => array('rgb' => 'b0cb96')
      ),
      'font' => array(
         'bold' => true,
         'name' => 'Arial',
         'size' => 10
      )
   )
);

// Header Data
$excel->setActiveSheetIndex(0)
   ->setCellValue('B5', 'Keterangan')
   ->setCellValue('C5', 'Data');

$excel->getActiveSheet()->getStyle('B5:C5')->applyFromArray(
   array(
      'fill' => array(
         'type' => PHPExcel_Style_Fill::FILL_SOLID,
         'color' => array('rgb' => 'd0d0d0')
      ),
      'font' => array(
         'bold' => true,
         'name' => 'Arial',
         'size' => 10
      )
   )
);

// Data Pribadi
$excel->setActiveSheetIndex(0)
   ->setCellValue('B6', 'DATA PRIBADI')
   ->setCellValue('B7', 'Nama lengkap')
   ->setCellValue('B8', 'Nama panggilan')
   ->setCellValue('B9', 'Telepon-1')
   ->setCellValue('B10', 'Telepon-2')
   ->setCellValue('B11', 'Email');

// Media Sosial
$excel->setActiveSheetIndex(0)
   ->setCellValue('B12', 'MEDIA SOSIAL')
   ->setCellValue('B13', 'Facebook')
   ->setCellValue('B14', 'Email facebook')
   ->setCellValue('B15', 'Instagram')
   ->setCellValue('B16', 'Twitter');

// Rencana Sewa
$excel->setActiveSheetIndex(0)
   ->setCellValue('B17', 'RENCANA SEWA')
   ->setCellValue('B18', 'Jenis alat')
   ->setCellValue('B19', 'Tanggal Register')
   ->setCellValue('B20', 'Acara')
   ->setCellValue('B21', 'Cabang Zenon');

// Data Penunjang
$excel->setActiveSheetIndex(0)
   ->setCellValue('B22', 'DATA PENUNJANG')
   ->setCellValue('B23', 'Mengetahui zenon dari mana?')
   ->setCellValue('B24', 'Sebelumnya sewa dimana?')
   ->setCellValue('B25', 'Atas nama siapa?');

// Keluarga yang Serumah
$excel->setActiveSheetIndex(0)
   ->setCellValue('B26', 'KELUARGA YG SERUMAH')
   ->setCellValue('B27', 'Atas nama siapa?')
   ->setCellValue('B28', 'Hubungan dengan penyewa')
   ->setCellValue('B29', 'Telepon (HP)')
   ->setCellValue('B30', 'Alamat');

// Pekerjaan
$excel->setActiveSheetIndex(0)
   ->setCellValue('B31', 'PEKERJAAN')
   ->setCellValue('B32', 'Pekerjaan sekarang')
   ->setCellValue('B33', 'Nama tempat kerja')
   ->setCellValue('B34', 'Jabatan kerja')
   ->setCellValue('B35', 'Alamat tempat kerja')
   ->setCellValue('B36', 'Telpon tempat kerja');

// Alamat tinggal sekarang
$excel->setActiveSheetIndex(0)
   ->setCellValue('B37', 'ALAMAT TINGGAL SEKARANG')
   ->setCellValue('B38', 'Status alamat tinggal sekarang')
   ->setCellValue('B39', 'Nama pemilik')
   ->setCellValue('B40', 'Telpon pemilik');

// Pendidikan
$excel->setActiveSheetIndex(0)
   ->setCellValue('B41', 'PENDIDIKAN')
   ->setCellValue('B42', 'Pendidikan sedang berjalan/terakhir')
   ->setCellValue('B43', 'Nama lembaga pendidikan')
   ->setCellValue('B44', 'Alamat lembaga pendidikan')
   ->setCellValue('B45', 'Info tambahan (angkatan masuk)');

// Verifikasi part
$excel->setActiveSheetIndex(0)
   ->setCellValue('E4', 'Verifikasi');

$excel->getActiveSheet()->getStyle('E4')->applyFromArray(
   array(
      'fill' => array(
         'type' => PHPExcel_Style_Fill::FILL_SOLID,
         'color' => array('rgb' => 'bed1f1')
      ),
      'font' => array(
         'bold' => true,
         'name' => 'Arial',
         'size' => 10
      )
   )
);

// Header Verif
$excel->setActiveSheetIndex(0)
   ->setCellValue('E5', 'PROS (+)')
   ->setCellValue('F5', 'CONS (-)');


//// Data Pribadi
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B6', 'DATA PRIBADI')
   //->setCellValue('B7', 'Nama lengkap')
   //->setCellValue('B8', 'Nama panggilan')
   //->setCellValue('B9', 'Telepon-1')
   //->setCellValue('B10', 'Telepon-2')
   //->setCellValue('B11', 'Email');

//// Media Sosial
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B12', 'MEDIA SOSIAL')
   //->setCellValue('B13', 'Facebook')
   //->setCellValue('B14', 'Email facebook')
   //->setCellValue('B15', 'Instagram')
   //->setCellValue('B16', 'Twitter');

//// Rencana Sewa
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B17', 'RENCANA SEWA')
   //->setCellValue('B18', 'Jenis alat')
   //->setCellValue('B19', 'Tanggal Register')
   //->setCellValue('B20', 'Acara')
   //->setCellValue('B21', 'Cabang Zenon');

//// Data Penunjang
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B22', 'DATA PENUNJANG')
   //->setCellValue('B23', 'Mengetahui zenon dari mana?')
   //->setCellValue('B24', 'Sebelumnya sewa dimana?')
   //->setCellValue('B25', 'Atas nama siapa?');

//// Keluarga yang Serumah
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B26', 'KELUARGA YG SERUMAH')
   //->setCellValue('B27', 'Atas nama siapa?')
   //->setCellValue('B28', 'Hubungan dengan penyewa')
   //->setCellValue('B29', 'Telepon (HP)')
   //->setCellValue('B30', 'Alamat');

//// Pekerjaan
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B31', 'PEKERJAAN')
   //->setCellValue('B32', 'Pekerjaan sekarang')
   //->setCellValue('B33', 'Nama tempat kerja')
   //->setCellValue('B34', 'Jabatan kerja')
   //->setCellValue('B35', 'Alamat tempat kerja')
   //->setCellValue('B36', 'Telpon tempat kerja');

//// Alamat tinggal sekarang
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B37', 'ALAMAT TINGGAL SEKARANG')
   //->setCellValue('B38', 'Status alamat tinggal sekarang')
   //->setCellValue('B39', 'Nama pemilik')
   //->setCellValue('B40', 'Telpon pemilik');

//// Pendidikan
//$excel->setActiveSheetIndex(0)
   //->setCellValue('B41', 'PENDIDIKAN')
   //->setCellValue('B42', 'Pendidikan sedang berjalan/terakhir')
   //->setCellValue('B43', 'Nama lembaga pendidikan')
   //->setCellValue('B44', 'Alamat lembaga pendidikan')
   //->setCellValue('B45', 'Info tambahan (angkatan masuk)');

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
