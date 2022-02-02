<?php
require_once 'PHPExcel/Classes/PHPExcel.php';
$archivo = "datos.xlsx";

$inputFileType = PHPExcel_IOFactory::identify($archivo);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objPHPExcel = $objReader->load($archivo);
$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();
$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn) - 1;

$num = 0; //Variable que sirve para mostrar las filas
$limCol = 0; //Queremos 2 columnas si y 2 no
$excel = array(array()); //Declaramos el array bidimensional 
$columnR = 0; //Columnas de las que sacamos datos
for ($row = 1; $row <= $highestRow; $row++) {
 for ($columnI = 0; $columnI <= $highestColumnIndex; $columnI++) {
  $limCol++;
  if ($limCol < 3) {
   $column = PHPExcel_Cell::stringFromColumnIndex($columnI);
   $valorCelda = mosCol($column, $row, $sheet);
   $excel[$row][$columnR] = $valorCelda;
   $columnR++;
  } else if ($limCol == 4) {
   $limCol = 0;
  }
 }

 $num++;
 $limCol = 0;
 $columnR = 0;
}

$excel2 = array(array()); //Declaramos el array bidimensional 
// Terminamos de formatear el excel
foreach ($excel as $f => $fila) {
 foreach ($fila as $c => $columna) {
  if ($c < 2 || $c % 2 == 1) {
   $excel2[$f][$c] = $columna;
  }
 }
}

function mosCol($column, $row, $sheet)
{
 $celda = $sheet->getCell($column . $row)->getValue();
 $celda = str_replace(",", "", $celda);
 $celda = str_replace("$", "", $celda);
 $celda = str_replace("(", "-", $celda);
 $celda = str_replace(")", "", $celda);
 $celda = trim($celda);
 return $celda;
}

// foreach ($excel2 as $f3 => $fila3) {
//  foreach ($fila3 as $c3 => $columna3) {
//   echo $f3 . "-" . $c3 . " -> " . $columna3 . " <br>";
//  }
//  echo "</br></br></br>";
// }
// die;
// create php excel object
$doc = new PHPExcel();

// set active sheet 
$doc->setActiveSheetIndex(0);

// read data to active sheet
$doc->getActiveSheet()->fromArray($excel2);

//save our workbook as this file name
$filename = 'orderFile.xlsx';
//mime type
header('Content-Type: application/vnd.ms-excel');
//tell browser what's the file name
header('Content-Disposition: attachment;filename="' . $filename . '"');

header('Cache-Control: max-age=0'); //no cache
//save it to Excel5 format (excel 2003 .XLS file), change this to 'Excel2007' (and adjust the filename extension, also the header mime type)
//if you want to save it as .XLSX Excel 2007 format

$objWriter = PHPExcel_IOFactory::createWriter($doc, 'Excel2007');

//force user to download the Excel file without writing it to server's HD
$objWriter->save('php://output');
