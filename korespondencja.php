<?php

    $cid = $_REQUEST['cid'];
    define('CID', $cid);
    define('READ_ONLY_SESSION',true);
    require_once('../../include.php');
    ModuleManager::load_modules();

    require_once 'modules/Libs/PHPExcel/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';

    $template_file = 'modules/driversRaport/theme/poczta.xls';
    $objReader = PHPExcel_IOFactory::createReader("Excel5");
    $objPHPExcel = $objReader->load($template_file);

    $rboCompany = new RBO_RecordsetAccessor('company');
    $records = $rboCompany->get_records(array('group' => array($_GET['commonDataKey']) ),array(),array());
    $row = 2;
    foreach($records as $record){
        $name = $record['company_name'];
        $address = $record['address_1'];
        $postalCode = $record['postal_code'];
        $postalCode = str_replace("-","",$postalCode);
        $city = $record['city'];
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0,$row,$name);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row,$address);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3,$row,$postalCode);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row,$city);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5,$row,'Polska');
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6,$row,'A');
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7,$row,'E');
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(8,$row,'0,25');
        $row += 1;
    } 
    
    $name = Utils_CommonDataCommon::get_array("Companies_Groups");
    $name  = __($name[$_GET['commonDataKey']]);
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel5");
    header('Content-Type: application/vnd.ms-excel');
    header("Content-Disposition: attachment; filename=\"$name.xls\"");
    header('Cache-Control: max-age=0');
    $objWriter->save('php://output');

    exit();