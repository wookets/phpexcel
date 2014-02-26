<?php

// Handles uploading and transform to json. 

ini_set("memory_limit", "256M");

require_once './Classes/PHPExcel/IOFactory.php';

if(!isset($_FILES['excelfile'])) {
	echo json_encode((object) array('status' => 'error', 'code' => 'InvalidParameter', 'message' => 'The excelfile parameter not found'));
	exit;
}

$inputFileName = $_FILES['excelfile']['tmp_name'];

/**  Identify the type of $inputFileName  **/
$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
/**  Create a new Reader of the type that has been identified  **/
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
if($inputFileType == 'CSV' || $inputFileType == 'csv') {
  ini_set("auto_detect_line_endings", true);
}
/**  Advise the Reader that we only want to load cell data  **/
//$objReader->setReadDataOnly(true); // We can't do this, because we want to read format data...
/**  Load $inputFileName to a PHPExcel Object  **/
$objPHPExcel = $objReader->load($inputFileName);

$excelArray = array();
foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
	$worksheetArray = array();
	//array_push($worksheetArray, $worksheet->getTitle());
	foreach ($worksheet->getRowIterator() as $row) {
		$rowArray = array();
		//array_push($excelArray, $row->getRowIndex());
		
		$cellIterator = $row->getCellIterator();
		$cellIterator->setIterateOnlyExistingCells(false); // Loop all cells, even if it is not set
		foreach ($cellIterator as $cell) {
			//if (!is_null($cell)) {
			//	array_push($excelArray, $cell->getCalculatedValue());
			//}
			try {
			  $dataType = $cell->getDataType();
			  // convert date...
			  $isDate = PHPExcel_Shared_Date::isDateTime($cell);
			  if($isDate) {
			    $dataType = 'date';
			    $calculatedValue = PHPExcel_Shared_Date::ExcelToPHPObject($cell->getValue())->format('Y-m-d');
			  } else {
  			  $calculatedValue = $cell->getCalculatedValue();
			  }
			  if($dataType == 's') $dataType = 'string';
			  if($dataType == 'n') $dataType = 'number';
			  $cellObj = array("column" => $cell->getColumn(), "row" => $cell->getRow(), "dataType" => $dataType,
			                   "value" => $cell->getValue(), "formattedValue" => $cell->getFormattedValue(), "calculatedValue" => $calculatedValue);
				array_push($rowArray, $cellObj);
			} catch(Exception $e) {
				array_push($rowArray, "NA");
			}
		}
		array_push($worksheetArray, $rowArray);
	}
	array_push($excelArray, $worksheetArray);
}

echo json_encode((object) array('status' => 'success', 'data' => $excelArray));

?>