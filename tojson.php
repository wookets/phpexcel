<?php

// Handles uploading and transform to json. 

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
/**  Advise the Reader that we only want to load cell data  **/
$objReader->setReadDataOnly(true);
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
			array_push($rowArray, $cell->getCalculatedValue());
		}
		array_push($worksheetArray, $rowArray);
	}
	array_push($excelArray, $worksheetArray);
}

echo json_encode((object) array('status' => 'success', 'data' => $excelArray));

?>