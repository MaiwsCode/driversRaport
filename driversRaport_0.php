<?php
defined("_VALID_ACCESS") || die('Direct access forbidden');

/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
class driversRaport extends Module {

    public function getCell($type){
        if($type == "bydło"){
           return 10;
        }
        else if ($type == "inny"){ // usluga transportowa
            return 12;
        }
        else if ($type == "service"){
            return 14;
        }
        else if ($type == "tucznik"){
            return 4;
        }
        else if ($type == "urlop"){
            return 15;
        }
        else if ($type == "warchlak"){
            return 6;
        }
        else if ($type == "brak_zlec"){
            return 16;
        }


    }



public function body(){
    require_once 'modules/Libs/PHPExcel/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';

    if(!$_REQUEST['selected_month']){
        $selected_month = date("m");
        $selected_year = date('Y');
        $selected_month = $selected_month-1;
        if($selected_month < 1){
            $selected_month = 12;
            $selected_year -= 1;
        }
    }else{
        $selected_month = $_REQUEST['selected_month'];
        $selected_year = $_REQUEST['year'];
    }
    $prev = $selected_month;
    $prev = $prev - 1;
    if($prev < 1){
        $prev = 12;
        $prev_year = $selected_year - 1;
    }else{
        $prev_year = $selected_year;
    }
    Base_ActionBarCommon::add(
        Base_ThemeCommon::get_template_file($this->get_type(), 'prev.png'),
        "Poprzedni miesiąc",
        $this->create_href ( array ('selected_month' => $prev, 'year' => $prev_year )),
        null,
        1
    );

    $next = $selected_month;
    $next += 1;
    if($next > 12){
        $next = 1;
        $next_year =  $selected_year + 1;
    }else{
        $next_year = $selected_year;
    }
    Base_ActionBarCommon::add(
        Base_ThemeCommon::get_template_file($this->get_type(), 'next.png'),
        "Następny miesiąc",
        $this->create_href ( array ('selected_month' => $next, 'year'=> $next_year )),
        null,
        2
    );
    Base_ActionBarCommon::add(
        'save',
        "Pobierz excel",
        $this->create_href ( array ('selected_month' => $selected_month, 'year' => $selected_year ,'download' => 'true')),
        null,
        3
    );

    $date_start = date("$selected_year-$selected_month-01");
    $date_end = date("$selected_year-".$selected_month."-t",strtotime($date_start));

    $redable_format = date("F",strtotime($date_start));
    print("Wybrany miesiąc - ".(__($redable_format))." $selected_year");
    $filename = date($selected_year."_".$selected_month."_01");
    $filename.=".xls";
    if($_REQUEST['download']){
        $rbo_drivers = new RBO_RecordsetAccessor('contact');
        $rbo_transports = new RBO_RecordsetAccessor("custom_agrohandel_transporty");
        $drivers = $rbo_drivers->get_records(array('group' => array('u_driver')),array(),array());
        $template_file = 'modules/driversRaport/theme/t1.xls';
        $objPHPExcel = new PHPExcel();
        $objReader = PHPExcel_IOFactory::createReader('Excel5');

        $objPHPExcel = $objReader->load($template_file);
        $objPHPExcel->setActiveSheetIndex(1);
        $pageIndex = 1;
        $driversSummary = array();
        foreach($drivers as $driver){
            $id = $driver->id;
            $workbook_page_name = "";
            $transports = $rbo_transports->get_records(array('driver_1' => $id,'>=date' => $date_start ,
                '<=date' => $date_end ),array(),array("date" => "ASC", "number" => "ASC"));

            $secondDriver = $rbo_transports->get_records(array('driver_2' => $id,'>=date' => $date_start ,
                '<=date' => $date_end ),array(),array("date" => "ASC", "number" => "ASC"));

            if($secondDriver != null || count($secondDriver) > 0){
                $transports += $secondDriver;
            }

            usort($transports, function ($item1, $item2) {
                return $item1['date'] <=> $item2['date'];
            });

            if($transports != null || count($transports) > 0){
                $workbook_page_name = $driver['last_name']." ".$driver['first_name'];
                $driverArrayIndex = $driver['last_name']."_".$driver['first_name'];
                $pageIndex = $pageIndex+1;
                $driversSummary[$driverArrayIndex][0] = $workbook_page_name;
                $objPHPExcel->setActiveSheetIndex(1);
                $newSheet = $objPHPExcel->getActiveSheet()->copy();
                $newSheet->setTitle($workbook_page_name);
                $objPHPExcel->addSheet($newSheet);
                $objPHPExcel->setActiveSheetIndex($pageIndex);
                $resize = $objPHPExcel->getActiveSheet();
                $resize->calculateColumnWidths(true);
                $objPHPExcel->getActiveSheet()->setCellValue("B1",$workbook_page_name);
                $objPHPExcel->getActiveSheet()->setCellValue("B2",__($redable_format)." ".$selected_year);
                $row = 5;
                $sums = array();
                $allLose = 0;
                $allTransported = 0;
                foreach($transports as $transport){
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0,$row, $transport['date']);
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,$row, $transport['number']);
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row,$transport['kmprzej']);
                    if($transport['driver_2'] == $id){
                        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3,$row,"X");
                    }

                    if($transport['kmprzej'] != ""){
                        $sums['allKm'] += $transport['kmprzej'];
                    }
                    $index = $this->getCell($transport['type']);
                    if($index < 14) {
                        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($index, $row, $transport['iloscrozl']);
                        $sums[$index] += $transport['iloscrozl'];
                        $driversSummary[$driverArrayIndex][$index] += 1;
                        $index += 1;
                        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($index, $row, $transport['iloscpadle']);
                        $sums[$index] += $transport['iloscpadle'];
                        $allLose += $transport['iloscpadle'];
                        $allTransported += $transport['iloscrozl'];
                        $driversSummary[$driverArrayIndex][1] += $transport['kmprzej'];
                    }else{
                        $driversSummary[$driverArrayIndex][$index] += 1;
                    }
                    for($x =0;$x<=13;$x++){
                        $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow($x,$row)->applyFromArray(array(
                            'borders' => array(
                                'allborders' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                                )
                            )));
                    }
                    $row = $row + 1;
                }
                $row = $row + 1;
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,$row,"SUMA: ");
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row, $sums['allKm']);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row, $sums[4]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5,$row, $sums[5]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6,$row, $sums[6]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7,$row, $sums[7]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(8,$row, $sums[8]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(9,$row, $sums[9]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(10,$row, $sums[10]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(11,$row, $sums[11]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(12,$row, $sums[12]);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(13,$row, $sums[13]);



                for($x =1;$x<=13;$x++){
                    $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow($x,$row)->applyFromArray(array(
                        'borders' => array(
                            'allborders' => array(
                                'style' => PHPExcel_Style_Border::BORDER_THIN
                            )
                        )));
                }

                $row = $row +1;
                $calc = $allLose / $allTransported;
                $calc = $calc * 100;
                $calc = round($calc, 4);
                $calc = str_replace(".",",",$calc);
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5,$row,$calc."%");
            }
        }
        $row = 5;
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setCellValue("B2",__($redable_format)." ".$selected_year);
        $allKm = 0;
        foreach($driversSummary as $driver){
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0,$row,$driver[0]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,$row,$driver[1]);
            $allKm += $driver[1];
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row,$driver[4]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3,$row,$driver[6]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row,$driver[10]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5,$row,$driver[14]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6,$row,$driver[12]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7,$row,$driver[15]);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(8,$row,$driver[16]);
            for($x =0;$x<=8;$x++){
                $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow($x,$row)->applyFromArray(array(
                    'borders' => array(
                    'allborders' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN)
                )));
            }
            $row = $row +  1;
        }
        $row = $row +  1;
        $objPHPExcel->removeSheetByIndex(1);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,$row,$allKm);
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('modules/driversRaport/theme/'.$filename);
        Epesi::redirect($_SERVER['document_root']."/modules/driversRaport/theme/".$filename);
        unlink($_SERVER['document_root']."/modules/driversRaport/theme/".$filename);
        }
    }
}