<?php
require_once "Classes/PHPExcel.php";
        $tmpfname = "tabula1.xlsx";
        $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
        $excelObj = $excelReader->load($tmpfname);
        $worksheet = $excelObj->getSheet(0);//
        $lastRow = $worksheet->getHighestRow();

        $data = [];
        for ($row = 1; $row <= $lastRow; $row++) {
             $data[] = [
                'A' => $worksheet->getCell('A'.$row)->getValue(),
                'B' => $worksheet->getCell('B'.$row)->getValue()
             ];
        }

echo json_encode($data);
?>