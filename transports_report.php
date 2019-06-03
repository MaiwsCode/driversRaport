<?php

$cid = $_REQUEST['cid'];
define('CID', $cid);
define('READ_ONLY_SESSION',true);

require_once('../../include.php');

ModuleManager::load_modules();

require_once 'modules/Libs/PHPExcel/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';

function color($excel,$range,$color){
    $excel->getActiveSheet()->getStyle($range)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB($color);
}

$template_file = 'modules/driversRaport/theme/transport_rozpiska.xls';
$objReader = PHPExcel_IOFactory::createReader("Excel5");
$objPHPExcel = $objReader->load($template_file);

$date = $_REQUEST['day'];
//$date = "2019-08-06";


$rboTransports = new RBO_RecordsetAccessor("custom_agrohandel_transporty");
$rboContacts = new RBO_RecordsetAccessor("contact");
$rboCompany = new RBO_RecordsetAccessor("company");
$rRboPurchasePlans = new RBO_RecordsetAccessor("custom_agrohandel_purchase_plans");
$rboVehicles = new RBO_RecordsetAccessor("custom_agrohandel_vehicle");
$transport['vehicle'];
$transports = $rboTransports->get_Records(array("date" => $date),array(),array());

$transportsCount = count($transports);

foreach($transports as $transport){
    if($transport['zakupy']){
        $copy = clone $objPHPExcel->getSheetByName("template");
        $copy->setTitle($transport['number']);
        $objPHPExcel->addSheet($copy);
    }
}
$page = 1;
foreach($transports as $transport){
    if($transport['zakupy']){
    //page
        $forceNull = false;
        $contact = $rboContacts->get_record($transport['driver_1']);
        $company = $rboCompany->get_record($transport['company']);
        $driver    = $contact['first_name']." ".$contact['last_name'];   //c4
        if($transport['driver_2']){
            $contact = $rboContacts->get_record($transport['driver_2']);
            $driver .= "/ ". $contact['first_name']." ".$contact['last_name'];
        }
        $objPHPExcel->setActiveSheetIndexByName($transport['number']);
        $objPHPExcel->getActiveSheet()->setCellValue("C1", $transport['number']);
        $objPHPExcel->getActiveSheet()->setCellValue("G1", str_replace("&quot;","", $company['company_name']) );
        $objPHPExcel->getActiveSheet()->setCellValue("C2", $date);
        $objPHPExcel->getActiveSheet()->setCellValue("C4", $driver);
        $car = $rboVehicles->get_record($transport['vehicle']);
        $objPHPExcel->getActiveSheet()->setCellValue("C6",$car['name'] );
        $objPHPExcel->getActiveSheet()->setCellValue("C7", $transport['kmprzej']);
        $row = 12;
        $weightOfAll = 0;
        foreach($transport['zakupy'] as $zakup){
            $zakup = $rRboPurchasePlans->get_record($zakup);
            //foreach 
            $objPHPExcel->getActiveSheet()->mergeCells("A$row:B$row");
            $objPHPExcel->getActiveSheet()->mergeCells("C$row:D$row");
            $objPHPExcel->getActiveSheet()->mergeCells("E$row:F$row");
            $objPHPExcel->getActiveSheet()->mergeCells("H$row:I$row");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, "Nr ubojowy:");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, "Klient:");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, $row, "Adres:");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, $row, "%");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7, $row, "Uwagi:");
            $style = [
                "font"=> [ 
                    "italic" => true,
                    "size" => 8
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    "wrap" => true,
                ],
                "fill" => [
                    'fillType' => PHPExcel_Style_Fill::FILL_SOLID,
                    'startColor' => [
                        'rgb' => 'd9d9d9',
                    ],
                ],
                
            ];
            $styleL = [
                "font"=> [ 
                    "italic" => true,
                    "size" => 8
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_BOTTOM,
                    "wrap" => true,
                ],
                "fill" => [
                    'fillType' => PHPExcel_Style_Fill::FILL_SOLID,
                    'startColor' => [
                        'rgb' => 'd9d9d9',
                    ],
                ],
            ];
            $nrStyle =  [
                "font"=> [ 
                    "bold" => true,
                    "size" => 11
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    "wrap" => true,
                ],
            ];
            $styleCC = [
                "font"=> [ 
                    "bold" => true,
                    "size" => 10
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    "wrap" => true,
                ],
            ];
            $styleCCS = [
                "font"=> [ 
                    "bold" => true,
                    "size" => 8
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    "wrap" => true,
                ],
            ];
            $opis = [
                "font"=> [ 
                    "bold" => true,
                    "size" => 8
                ],
                "alignment" => [
                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                    "vertical" => PHPExcel_Style_Alignment::VERTICAL_TOP,
                    "wrap" => true,
                ],
            ];



            color($objPHPExcel,"A$row:I$row","d3d3d3");
            $objPHPExcel->getActiveSheet()->getStyle("A$row:I$row")->applyFromArray($style);

            $row2 = $row + 2;
            $borders = [
                "borders" => [
                    "allborders" => [
                        "style" => PHPExcel_Style_Border::BORDER_HAIR
                    ]
                ]
            ];
            $objPHPExcel->getActiveSheet()->getStyle("A$row:I$row2")->applyFromArray($borders);
            //dane 
            $row +=1;
            $companyRecord = $rboCompany->get_record($zakup['company']);
            $clientText = $companyRecord->get_val("parent_company",$nolink = true)."\n";
            $clientText .= $companyRecord->get_val("company_name",$nolink = true); 
            $clientText = str_replace("&nbsp;", " ", $clientText);
            $clientText = str_replace("&quot;", " ", $clientText);

            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, $zakup['numer_ubojowy']);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, $clientText);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, $row, $companyRecord['address_1']." ".$companyRecord['postal_code']." ".$companyRecord['city']);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, $row, $zakup['deduction']);
            $note = "";
            if($transport['noterozl']){
                $note .= $transport['noterozl'].". ";
            }
            if($zakup['noteh']){
                $note .= $zakup['noteh'].". ";
            }
            if($zakup['note']){
                $note .= $zakup['note'].". ";
            }
            if($zakup['notek']){
                $note .= $zakup['notek'].". ";
            }
            if($zakup['additional_fixing']){
                $note .= $zakup['additional_fixing'].". ";
            }
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7, $row, $note );

            $nextRow = $row + 1;
            $contact = $rboContacts->get_record($companyRecord['account_manager']);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $nextRow, "Opiekun:");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3, $nextRow, $contact['first_name']);

            $lines_arr = preg_split('/\n|\r/',$note);
            $num_newlines = count($lines_arr); 

            $objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(50.50 + $num_newlines * 15);
            $objPHPExcel->getActiveSheet()->mergeCells("A$row:B$nextRow");
            $objPHPExcel->getActiveSheet()->mergeCells("C$row:D$row");
            $objPHPExcel->getActiveSheet()->mergeCells("E$row:F$nextRow");
            $objPHPExcel->getActiveSheet()->mergeCells("G$row:G$nextRow");
            $objPHPExcel->getActiveSheet()->mergeCells("H$row:I$nextRow");

            $objPHPExcel->getActiveSheet()->getStyle("A$row:B$nextRow")->applyFromArray($nrStyle);
            $objPHPExcel->getActiveSheet()->getStyle("C$row:D$row")->applyFromArray($styleCC);
            $objPHPExcel->getActiveSheet()->getStyle("E$row:F$nextRow")->applyFromArray($styleCC);
            $objPHPExcel->getActiveSheet()->getStyle("G$row:G$nextRow")->applyFromArray($styleCC);
            $objPHPExcel->getActiveSheet()->getStyle("H$row:I$nextRow")->applyFromArray($opis);


            $objPHPExcel->getActiveSheet()->getStyle("C$nextRow")->applyFromArray($styleL);
            $objPHPExcel->getActiveSheet()->getStyle("D$nextRow")->applyFromArray($styleCCS);

            //opiekun handlowiec = index +1
            $row += 1;
            //dane 
            //druga linia
            $row += 1;
            $objPHPExcel->getActiveSheet()->getStyle("A$row:G$row")->applyFromArray($style);
            color($objPHPExcel,"A$row:G$row","d3d3d3");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, "Ilość");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3, $row, "Waga");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, $row, "Warunek");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5, $row, "Padłe");
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, $row, "Śr. waga");
            $row2 = $row + 1;
            $objPHPExcel->getActiveSheet()->getStyle("A$row:G$row2")->applyFromArray($borders);
            $bordersLineTop = [
                "borders" => [
                    "top" => [
                        "style" => PHPExcel_Style_Border::BORDER_THIN
                    ]
                ]
            ];
            $bordersLineBot = [
                "borders" => [
                    "bottom" => [
                        "style" => PHPExcel_Style_Border::BORDER_THIN
                    ]
                ]
            ];
            $bordersLineLeft = [
                "borders" => [
                    "left" => [
                        "style" => PHPExcel_Style_Border::BORDER_THIN
                    ]
                ]
            ];
            $bordersLineRight = [
                "borders" => [
                    "right" => [
                        "style" => PHPExcel_Style_Border::BORDER_THIN
                    ]
                ]
            ];
            $objPHPExcel->getActiveSheet()->getStyle("A$row:I$row")->applyFromArray($bordersLineTop);
            $objPHPExcel->getActiveSheet()->getStyle("A$row:G$row")->applyFromArray($bordersLineBot);

            $row2 = $row +1;
            $objPHPExcel->getActiveSheet()->getStyle("G$row")->applyFromArray($bordersLineRight);
            $objPHPExcel->getActiveSheet()->getStyle("G$row2")->applyFromArray($bordersLineRight);
            $objPHPExcel->getActiveSheet()->getStyle("A$row2:G$row2")->applyFromArray($bordersLineBot);
            $start = $row - 2;
            $end = $row + 1;
            $objPHPExcel->getActiveSheet()->getStyle("A$start:A$end")->applyFromArray($bordersLineLeft);
            $objPHPExcel->getActiveSheet()->getStyle("B$start:B$end")->applyFromArray($bordersLineRight);
            $start = $row - 3;
            $objPHPExcel->getActiveSheet()->getStyle("A$start:I$start")->applyFromArray($bordersLineTop);
            $objPHPExcel->getActiveSheet()->getStyle("A$start:I$start")->applyFromArray($bordersLineBot);
            $end = $start + 2;
            $objPHPExcel->getActiveSheet()->getStyle("I$start:I$end")->applyFromArray($bordersLineRight);

            $row += 1;
            $row2 = $row - 1;
            $objPHPExcel->getActiveSheet()->mergeCells("A$row2:B$row");
            
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row2, "Załadowano");

            $full = $zakup['wagazalak'];
            $empty = $zakup['wagazala'];

            $weight = 0;
            $avgWeight = 0;
            
            if($empty == 0 || strlen($empty) == 0 ){
                $weight = $full;
            }else{
                $weight = $full - $empty;
            }
            if(!$weight){
                $weight = 0;
            }
            $szt = $zakup['sztukzal'];
            $weightOfAll += $weight;
            $avgWeight = $weight / $zakup['sztukzal'];
            $avgWeight = round($avgWeight,2);
            $warunki = $zakup['warunkizal'];
            if(!$warunki){
                $warunki = 0;
            }
            $upadki = $zakup['upadrozl'];
            if(!$upadki){
                $upadki = 0;
            }
            if(!$avgWeight){
                $avgWeight = 0;
            }
            if(!$szt){
                $szt = 0;
            }
            if(!$zakup['sztukzal']){
                $avgWeight = 0;
                $szt = 0;
                $forceNull = true;
            }
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, $szt);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3, $row, $weight);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, $row, $warunki );
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5, $row,$upadki);
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, $row, $avgWeight);
            $objPHPExcel->getActiveSheet()->getStyle("A$row")->applyFromArray($styleCC);
            $objPHPExcel->getActiveSheet()->getStyle("C$row:G$row")->applyFromArray($styleL);

            color($objPHPExcel,"A$row","d3d3d3");
            $row += 2;
            //nowy blok

            //end foreach
        }
        $wagaPrzed = $transport['wagarozprzed'];
        $wagaPo = $transport['wagarozpo'];
        if(!$wagaPo){
            $wagaPo = 0;
        }
        if(!$wagaPrzed){
            $wagaPrzed = 0;
        }

        $trWeight = 0;
        if(($wagaPo == 0 || strlen($wagaPo) == 0) && $transport['iloscrozl'] > 0 ){
            $weightOfAll = $weightOfAll - $wagaPrzed;
            $weightOfAll = $weightOfAll / $transport['iloscrozl'];
            $trWeight = $wagaPrzed;
        }else{
            $weightOfAll = $weightOfAll - ($wagaPrzed - $wagaPo);
            $trWeight = $wagaPrzed - $wagaPo;
            $weightOfAll = $weightOfAll / $transport['iloscrozl'];
        }
        if($weightOfAll < 0){
            $weightOfAll = "";
        }
        $weightOfAll = round($weightOfAll,2);
        if(!$transport['iloscpadle']){
            $transport['iloscpadle'] = 0;
        }
        if(!$transport['iloscwaru']){
            $transport['iloscwaru']  = 0 ;
        }
        if($forceNull){
            $weightOfAll = 0;
        }
        if($wagaPo == 0 && $wagaPrzed == 0){
            $weightOfAll = 0;
        }

        $objPHPExcel->getActiveSheet()->getStyle("C10:I10")->applyFromArray($styleCC);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, 10, $transport['iloscrozl']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, 10, $trWeight);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, 10, $transport['iloscpadle']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7, 10, $transport['iloscwaru']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(8, 10, $weightOfAll);
        $page += 1;
    }
}
$objPHPExcel->setActiveSheetIndexByName("template");
$objPHPExcel->removeSheetByIndex($objPHPExcel->getActiveSheetIndex());

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel5");
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment; filename="rozpiska-'.$date.'".xls"');
header('Cache-Control: max-age=0');
$objWriter->save('php://output');
exit();
