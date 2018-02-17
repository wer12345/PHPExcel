<?php
require_once './Classes/PHPExcel.php';
require_once './function.php';

// Create PHPExcel object
$excel = new PHPExcel();

// remove gridlines
$excel->getActiveSheet()->setShowGridlines(false);

cellStyle($excel, 'B6', $title);
cellStyle($excel, 'B12', $title);
cellStyle($excel, 'B17', $title);
cellStyle($excel, 'B22', $title);
cellStyle($excel, 'B26', $title);
cellStyle($excel, 'B31', $title);
cellStyle($excel, 'B37', $title);
cellStyle($excel, 'B41', $title);
cellStyle($excel, 'E4', $title);
cellStyle($excel, 'E5:F5', $title);
cellStyle($excel, 'E6:F6', $title);
cellStyle($excel, 'E14', $title);
cellStyle($excel, 'E28', $title);
cellStyle($excel, 'E36', $title);
cellStyle($excel, 'E44', $title);
cellStyle($excel, 'H5:I5', $title);

// Set Column Width
cellWidth($excel, 'A', 2);
cellWidth($excel, 'B', 35);
cellWidth($excel, 'C', 35);
cellWidth($excel, 'D', 4);
cellWidth($excel, 'E', 40);
cellWidth($excel, 'F', 40);
cellWidth($excel, 'G', 4);
cellWidth($excel, 'H', 20);
cellWidth($excel, 'I', 40);

// cell alignment
//$excel->getActiveSheet()->getStyle('B2:C45')->applyFromArray(
//array(
//'alignment' => array(
//'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
//'hrizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
//)
//)
//);

$excel->getActiveSheet()->getStyle('B2:C45')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$excel->getActiveSheet()->getStyle('h6:I49')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

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
$excel->getActiveSheet()->mergeCells('E6:F6');
$excel->getActiveSheet()->mergeCells('E14:F14');
$excel->getActiveSheet()->mergeCells('E28:F28');
$excel->getActiveSheet()->mergeCells('E36:F36');
$excel->getActiveSheet()->mergeCells('E44:F44');
$excel->getActiveSheet()->mergeCells('H6:H8');
$excel->getActiveSheet()->mergeCells('H9:H11');
$excel->getActiveSheet()->mergeCells('H12:H14');
$excel->getActiveSheet()->mergeCells('H15:H17');
$excel->getActiveSheet()->mergeCells('I6:I8');
$excel->getActiveSheet()->mergeCells('I9:I11');
$excel->getActiveSheet()->mergeCells('I12:I14');
$excel->getActiveSheet()->mergeCells('I15:I17');
$excel->getActiveSheet()->mergeCells('H27:I27');
$excel->getActiveSheet()->mergeCells('H28:H37');
$excel->getActiveSheet()->mergeCells('H38:H45');
$excel->getActiveSheet()->mergeCells('H46:H52');
$excel->getActiveSheet()->mergeCells('I28:I37');
$excel->getActiveSheet()->mergeCells('I38:I45');
$excel->getActiveSheet()->mergeCells('I46:I52');


// give border

$boderStyle = array(
   'borders' => array(
      'allborders' => array(
         'style' => PHPExcel_Style_Border::BORDER_THIN,
         'color' => array('rgb' => 'a6a6a6')
      )
   )
);

$fontStyle = array(
   'font' => array(
      'size' => 10,
      'name' => 'Arial'

   )
);

$excel->getActiveSheet()->getStyle('B5:C45')->applyFromArray($boderStyle);
$excel->getActiveSheet()->getStyle('H20:I24')->applyFromArray($boderStyle);

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

$excel->getActiveSheet()->getStyle('E5:F45')->applyFromArray(
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

$excel->getActiveSheet()->getStyle('H6:I17')->applyFromArray(
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

//$excel->getActiveSheet()->getStyle('H20:I24')->applyFromArray(
//array(
//'borders' => array(
//'outline' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//),
//'vertical' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//),
//'inside' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//)
//),
//'font' => array(
//'bold' => true,
//'name' => 'Arial',
//'size' => 10
//)
//)
//);

//$excel->getActiveSheet()->getStyle('H27:I52')->applyFromArray(
//array(
//'borders' => array(
//'outline' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//),
//'vertical' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//),
//'inside' => array(
//'style' => PHPExcel_Style_Border::BORDER_THIN,
//'color' => array('rgb' => 'a6a6a6')
//)
//),
//'font' => array(
//'bold' => true,
//'name' => 'Arial',
//'size' => 10
//)
//)
//);

//$excel->getActiveSheet()->getStyle('H27:I52')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN)->getColor()->setRGB('a6a6a6');
//$excel->getActiveSheet()->getStyle('H20:I24')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN)->getColor()->setRGB('a6a6a6');

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
      ),
      'alignment' => array(
         'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
      )
   )
);

// Header Verif
$excel->setActiveSheetIndex(0)
   ->setCellValue('E5', 'PROS (+)')
   ->setCellValue('F5', 'CONS (-)');


// Umum
$excel->setActiveSheetIndex(0)
   ->setCellValue('E6', '## UMUM')
   ->setCellValue('E7', '-Cek daftar blacklist:')
   ->setCellValue('E8', '-Cek validasi NIK KTP (Via App dan web KPU):')
   ->setCellValue('E9', '-Cek keamanan data (semua dokumen):')
   ->setCellValue('E10', '-Tracking nama:')
   ->setCellValue('E11', '-Tracking HP:')
   ->setCellValue('E12', '-Apakah pemilik sebuah vendor:')
   ->setCellValue('E13', '');

// Medsos Fb
$excel->setActiveSheetIndex(0)
   ->setCellValue('E14', '## MEDSOS FB')
   ->setCellValue('E15', '-dp muka:')
   ->setCellValue('E16', '-foto selfie:')
   ->setCellValue('E17', '-awal bikin:')
   ->setCellValue('E18', '-LU:')
   ->setCellValue('E19', '-Interval post:')
   ->setCellValue('E20', '-valid bday:')
   ->setCellValue('E21', '-valid kerjaan:')
   ->setCellValue('E22', '-valid sekolah:')
   ->setCellValue('E23', '-valid hub, suami istri:')
   ->setCellValue('E24', '-portfolio:')
   ->setCellValue('E25', '-mf:')
   ->setCellValue('E26', '-lainnya:')
   ->setCellValue('E27', '');

// Medsos Ig
$excel->setActiveSheetIndex(0)
   ->setCellValue('E28', '## MEDSOS IG')
   ->setCellValue('E29', '-dp muka:')
   ->setCellValue('E30', '-post:')
   ->setCellValue('E31', '-follower:')
   ->setCellValue('E32', '-portfolio:')
   ->setCellValue('E33', '-selfie:')
   ->setCellValue('E34', '-LU:')
   ->setCellValue('E35', '');

// Medsos Tw
$excel->setActiveSheetIndex(0)
   ->setCellValue('E36', '## MEDSOS TW')
   ->setCellValue('E37', '-dp muka:')
   ->setCellValue('E38', '-join:')
   ->setCellValue('E39', '-post:')
   ->setCellValue('E40', '-follower:')
   ->setCellValue('E41', '-selfie:')
   ->setCellValue('E42', '-LU:')
   ->setCellValue('E43', '');

// Website
$excel->setActiveSheetIndex(0)
   ->setCellValue('E36', '## WEBSITE')
   ->setCellValue('E37', 'website pribadi/vendor:')
   ->setCellValue('E38', 'whois:');

// Header Note
$excel->setActiveSheetIndex(0)
   ->setCellValue('H5', 'NOTE')
   ->setCellValue('I5', 'KETERANGAN');

$excel->setActiveSheetIndex(0)
   ->setCellValue('H6', 'KEKURANGAN DATA:')
   ->setCellValue('H9', 'KESIMPULAN CCO:')
   ->setCellValue('H12', 'KESIMPULAN SPV:')
   ->setCellValue('H15', 'KESIMPULAN CCO:');

$excel->setActiveSheetIndex(0)
   ->setCellValue('H20', 'LEVEL CCO')
   ->setCellValue('H21', 'LEVEL SPV:')
   ->setCellValue('H22', 'LEVEL GM')
   ->setCellValue('H23', 'LEVEL AKHIR')
   ->setCellValue('H24', 'STATUS');

$excel->setActiveSheetIndex(0)
   ->setCellValue('H27', 'UPLOAD FOTO & DOKUMEN')
   ->setCellValue('H28', 'Foto selfie terbaru')
   ->setCellValue('H38', 'KTP (Wajib)')
   ->setCellValue('H46', 'Dokumen berharga lain');

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
