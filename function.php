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

// cell input data
function data($excel, $data = '')
{
   // Tanggal Form Registration
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C2', 'Tanggal Register')
      ->setCellValue('C3', 'Tanggal Verif');

   // Bagian Data Pribadi
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C7', "Nama Lengkap")
      ->setCellValue('C8', "Name Panggilan")
      ->setCellValue('C9', "Telepon-1")
      ->setCellValue('C10', "Telepon-2")
      ->setCellValue('C11', "Email");

  // Media Sosial 
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C13', "Facebok")
      ->setCellValue('C14', "Email Facebook")
      ->setCellValue('C15', "Instagram")
      ->setCellValue('C16', "Twitter");

   // Rencana Sewa
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C18', 'Jenis alat')
      ->setCellValue('C19', 'Tanggal Register')
      ->setCellValue('C20', 'Acara')
      ->setCellValue('C21', 'Cabang Zenon');

   // Data Penunjang
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C23', 'Mengetahui zenon dari mana?')
      ->setCellValue('C24', 'Sebelumnya sewa dimana?')
      ->setCellValue('C25', 'Atas nama siapa?');
   
   // Keluarga yang Serumah
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C27', 'Atas nama siapa?')
      ->setCellValue('C28', 'Hubungan dengan penyewa')
      ->setCellValue('C29', 'Telepon (HP)')
      ->setCellValue('C30', 'Alamat');

   // Pekerjaan
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C32', 'Pekerjaan sekarang')
      ->setCellValue('C33', 'Nama tempat kerja')
      ->setCellValue('C34', 'Jabatan kerja')
      ->setCellValue('C35', 'Alamat tempat kerja')
      ->setCellValue('C36', 'Telpon tempat kerja');

   // Alamat tinggal sekarang
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C38', 'Status alamat tinggal sekarang')
      ->setCellValue('C39', 'Nama pemilik')
      ->setCellValue('C40', 'Telpon pemilik');

   // Pendidikan
   $excel->setActiveSheetIndex(0)
      ->setCellValue('C42', 'Pendidikan sedang berjalan/terakhir')
      ->setCellValue('C43', 'Nama lembaga pendidikan')
      ->setCellValue('C44', 'Alamat lembaga pendidikan')
      ->setCellValue('C45', 'Info tambahan (angkatan masuk)');
   
   // Umum
   $excel->setActiveSheetIndex(0)
      ->setCellValue('F7', '-Cek daftar blacklist:')
      ->setCellValue('F8', '-Cek validasi NIK KTP (Via App dan web KPU):')
      ->setCellValue('F9', '-Cek keamanan data (semua dokumen):')
      ->setCellValue('F10', '-Tracking nama:')
      ->setCellValue('F11', '-Tracking HP:')
      ->setCellValue('F12', '-Apakah pemilik sebuah vendor:')
      ->setCellValue('F13', '');

   // Medsos Fb
   $excel->setActiveSheetIndex(0)
      ->setCellValue('F15', '-dp muka:')
      ->setCellValue('F16', '-foto selfie:')
      ->setCellValue('F17', '-awal bikin:')
      ->setCellValue('F18', '-LU:')
      ->setCellValue('F19', '-Interval post:')
      ->setCellValue('F20', '-valid bday:')
      ->setCellValue('F21', '-valid kerjaan:')
      ->setCellValue('F22', '-valid sekolah:')
      ->setCellValue('F23', '-valid hub, suami istri:')
      ->setCellValue('F24', '-portfolio:')
      ->setCellValue('F25', '-mf:')
      ->setCellValue('F26', '-lainnya:')
      ->setCellValue('F27', '');

   // Medsos Ig
   $excel->setActiveSheetIndex(0)
      ->setCellValue('F29', '-dp muka:')
      ->setCellValue('F30', '-post:')
      ->setCellValue('F31', '-follower:')
      ->setCellValue('F32', '-portfolio:')
      ->setCellValue('F33', '-selfie:')
      ->setCellValue('F34', '-LU:')
      ->setCellValue('F35', '');

   // Medsos Tw
   $excel->setActiveSheetIndex(0)
      ->setCellValue('F37', '-dp muka:')
      ->setCellValue('F38', '-join:')
      ->setCellValue('F39', '-post:')
      ->setCellValue('F40', '-follower:')
      ->setCellValue('F41', '-selfie:')
      ->setCellValue('F42', '-LU:')
      ->setCellValue('F43', '');

   // Website
   $excel->setActiveSheetIndex(0)
      ->setCellValue('F45', 'website pribadi/vendor:')
      ->setCellValue('F46', 'whois:');


   $excel->setActiveSheetIndex(0)
      ->setCellValue('I6', 'KEKURANGAN DATA:')
      ->setCellValue('I9', 'KESIMPULAN CCO:')
      ->setCellValue('I12', 'KESIMPULAN SPV:')
      ->setCellValue('I15', 'KESIMPULAN CCO:');

   $excel->setActiveSheetIndex(0)
      ->setCellValue('I20', 'LEVEL CCO')
      ->setCellValue('I21', 'LEVEL SPV:')
      ->setCellValue('I22', 'LEVEL GM')
      ->setCellValue('I23', 'LEVEL AKHIR')
      ->setCellValue('I24', 'STATUS');

   $excel->setActiveSheetIndex(0)
      ->setCellValue('I27', 'UPLOAD FOTO & DOKUMEN')
      ->setCellValue('I28', 'Foto selfie terbaru')
      ->setCellValue('I38', 'KTP (Wajib)')
      ->setCellValue('I46', 'Dokumen berharga lain');
}


// cell title
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


// cell border
$borderStyle = array(
   'borders' => array(
      'allborders' => array(
         'style' => PHPExcel_Style_Border::BORDER_THIN,
         'color' => array('rgb' => 'a6a6a6')
      )
   )
);

// cell font
$fontStyle = array(
   'font' => array(
      'size' => 10,
      'name' => 'Arial'

   )
);

// cell background
$fillStyle = array(
   'fill' => array(
      'type' => PHPExcel_Style_Fill::FILL_SOLID, 
      'color' => array('rgb' => 'f3bc93')
   ),
   'font' => array(
      'bold' => true
   )
);

$fillStyle2 = array(
   'fill' => array(
      'type' => PHPExcel_Style_Fill::FILL_SOLID, 
      'color' => array('rgb' => 'fedd8b')
   ),
   'font' => array(
      'bold' => true
   )
);

$fillStyle3 = array(
   'fill' => array(
      'type' => PHPExcel_Style_Fill::FILL_SOLID, 
      'color' => array('rgb' => 'ecc0c1')
   ),
   'font' => array(
      'bold' => true
   )
);

$fillStyle4 = array(
   'fill' => array(
      'type' => PHPExcel_Style_Fill::FILL_SOLID, 
      'color' => array('rgb' => 'b0cb96')
   ),
   'font' => array(
      'bold' => true
   )
);

?>
