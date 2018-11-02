<?php
defined("_VALID_ACCESS") || die('Direct access forbidden');

/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
class driversRaport extends Module { 

public function body(){
    require_once 'modules/Libs/PHPExcel/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';

    if(!$_REQUEST['selected_month']){
        $selected_month = date("m");
        $selected_month = $selected_month-1;
    }else{
        $selected_month = $_REQUEST['selected_month'];
    }

    Base_ActionBarCommon::add(
        Base_ThemeCommon::get_template_file($this->get_type(), 'prev.png'),
        "Poprzedni miesiąc",
        $this->create_href ( array ('selected_month' => ($selected_month -1))),
        null,
        1
    );
    Base_ActionBarCommon::add(
        Base_ThemeCommon::get_template_file($this->get_type(), 'next.png'),
        "Następny miesiąc",
        $this->create_href ( array ('selected_month' => ($selected_month + 1))),
        null,
        2
    );
    Base_ActionBarCommon::add(
        'save',
        "Pobierz excel",
        $this->create_href ( array ('selected_month' => ($selected_month),'download' => 'true')),
        null,
        3
    );

    $date_start = date("Y-".$selected_month."-01");
    $date_end = date("Y-".$selected_month."-t",strtotime($date_start));
 
    $redable_format = date("F",strtotime($date_start));
    print("Wybrany miesiąc - ".(__($redable_format)));
    $filename = date("Y_".$selected_month."_01");
    $filename.=".xls";
    if($_REQUEST['download']){
        $rbo_drivers = new RBO_RecordsetAccessor('contact');
        $rbo_transports = new RBO_RecordsetAccessor("custom_agrohandel_transporty"); 
        $drivers = $rbo_drivers->get_records(array('group' => array('u_driver')),array(),array());


        $template_file = 'modules/driversRaport/theme/t1.xls';
        $objPHPExcel = new PHPExcel();
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        
        $objPHPExcel = $objReader->load($template_file);
        $objPHPExcel->setActiveSheetIndex(0);

        $d = 0;
        foreach($drivers as $driver){
            $id = $driver->id;
            $workbook_page_name = "";
            $transports = $rbo_transports->get_records(array('driver_1' => $id,'>=date' => $date_start ,
            '<=date' => $date_end, '>iloscrozl' => 0  ),array(),array("date" => "ASC", "number" => "ASC"));
            if($transports != null || count($transports) > 0){
                $workbook_page_name = $driver['last_name']." ".$driver['first_name'];
                $d = $d+1;
                $objPHPExcel->setActiveSheetIndex(0);
                $newSheet = $objPHPExcel->getActiveSheet()->copy();
                $newSheet->setTitle($workbook_page_name);
                $objPHPExcel->addSheet($newSheet);
                $objPHPExcel->setActiveSheetIndex($d);
                $resize = $objPHPExcel->getActiveSheet();
                $resize->calculateColumnWidths(true);
                $objPHPExcel->getActiveSheet()->setCellValue("B1",$workbook_page_name);

                $objPHPExcel->getActiveSheet()->setCellValue("B2",__($redable_format)." ".date("Y"));
                $row = 5;
                $sum_km = 0;
                $sum_drop = 0;
                $sum_lose = 0;
                foreach($transports as $transport){
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0,$row, $transport['date']);
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,$row, $transport['number']);
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row,$transport['kmprzej']);
                    if($transport['kmprzej'] != ""){
                        $sum_km += $transport['kmprzej'];
                    }
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3,$row, $transport['iloscrozl']);
                    if($transport['iloscrozl'] != ""){
                        $sum_drop += $transport['iloscrozl'];
                    }
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row,$transport['iloscpadle']);
                    if($transport['iloscpadle'] != ""){
                        $sum_lose += $transport['iloscpadle'];
                    }
                    for($x =0;$x<=4;$x++){
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
                $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow(1,$row)->getFont()->applyFromArray(array('name'      => 'Calibri','bold'      => true,'italic'    => false,'underline' => false,'strike'    => false,'color'     => array('rgb' => '000000')));
                
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2,$row,$sum_km);
                $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow(2,$row)->getFont()->applyFromArray(array('name'      => 'Calibri','bold'      => true,'italic'    => false,'underline' => false,'strike'    => false,'color'     => array('rgb' => '000000')));

                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3,$row,$sum_drop);
                $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow(3,$row)->getFont()->applyFromArray(array('name'      => 'Calibri','bold'      => true,'italic'    => false,'underline' => false,'strike'    => false,'color'     => array('rgb' => '000000')));

                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row,$sum_lose);
                $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow(4,$row)->getFont()->applyFromArray(array('name'      => 'Calibri','bold'      => true,'italic'    => false,'underline' => false,'strike'    => false,'color'     => array('rgb' => '000000')));

                $row = $row +1;
                $calc = $sum_lose / $sum_drop;
                $calc = round($calc, 4);
                $calc = str_replace(".",",",$calc);
                $cell = "E".$row."";
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4,$row,$calc."%");

            }
        }
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('modules/driversRaport/theme/'.$filename);

        Epesi::redirect($_SERVER['document_root']."/modules/driversRaport/theme/".$filename);
        unlink($_SERVER['document_root']."/modules/driversRaport/theme/".$filename);
        }
    } 
}
