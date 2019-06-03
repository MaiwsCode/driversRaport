<?php
defined("_VALID_ACCESS") || die('Direct access forbidden');

/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
class driversRaport extends Module {

public function body(){

    $tabbed_browser = $this->init_module('Utils/TabbedBrowser');
    $tabbed_browser->set_tab(__('Raport kierowców'), array($this, 'main'));
    $tabbed_browser->set_tab(__('Poczta'), array($this, 'postal'));
    $tabbed_browser->set_tab(__('Raport transportowy'), array($this, 'transport'));

    $this->display_module($tabbed_browser);
}
public function transport(){
    Base_ThemeCommon::install_default_theme($this->get_type());
    $theme = $this->init_module('Base/Theme');
        $theme->display('transport');
        $form = $this->init_module('Libs/QuickForm');
        $form->addElement('datepicker', 'from', 'Od');
        $form->addElement('datepicker', 'to', 'Do');
        $form->addElement('text', 'zaokr', 'Zaokrąglenie');
        $form->addElement("submit", "submit", "Porównaj");
        $form->display();
        $prec = 100;
        $from = date('Y-m-d', strtotime('-3 month'));
        $to = date('Y-m-t');
        if($form->validate()){
            $values = $form->exportValues();
            if(count($values['zaokr']) && $values['zaokr'] != '' ){
                $prec = $values['zaokr'];
            }
            if(count($values['from']) && $values['from'] != '' ){
                $from = $values['from'];
            }
            if(count($values['to']) && $values['to'] != '' ){
                $to = $values['to'];
            }
        }
        print("OKRES OD: ".$from." DO: ".$to);
        $title = "Raport transportów";
        $rboTransports = new RBO_RecordsetAccessor("custom_agrohandel_transporty");
        $companes = new RBO_RecordsetAccessor("company");
        $bought = new RBO_RecordsetAccessor("custom_agrohandel_purchase_plans");
        $transports =  $rboTransports->get_records(array('>=date' => $from, '<=date' => $to),array(),array('date' => "ASC"));  
        $dataToParse = [];
        foreach($transports as $transport){
            $is_ubojnia = $companes->get_record($transport["company"]);
            if(!$is_ubojnia['group']['baza_tr']){
                $zakupy = $transport['zakupy'];
                foreach($zakupy as $zakup){
                    // suma z dnia poprzez zapupy przypiete pod tranport
                    $record = $bought->get_record($zakup);
                    $szt = $record['sztukzal'];
                    $dataToParse[$transport['date']]['date'] = $transport['date']; 
                    $dataToParse[$transport['date']]['szt'] +=  $szt;
                }
            }
        }
        $dataOutput = array();
        foreach($dataToParse as $transport){
            $szt = $transport['szt'];
            $szt = $szt / $prec;
            $szt = round($szt,0);
            $szt = $szt * $prec; 
            $dataOutput[$szt]['szt'] = $szt;
            if($dataOutput[$szt]['date'] != $transport['date']){
                $dataOutput[$szt]['date'] = $transport['date'];
                $dataOutput[$szt]['count'] += 1;
            }
        }
        $data = array();
        foreach($dataOutput as $transport){
            $data[] = array("y" => $transport['count'], "x" => $transport['szt']);
        }

        $data = (json_encode($data));
        $type = "column"; // bar, line, area, pie, etc
        load_js($this->get_module_dir()."js/canvasjs.min.js");
        eval_js('
        jq(document).ready(function (){ 
            var chart = new CanvasJS.Chart("chartContainer", {
                animationEnabled: true,
                exportEnabled: true,
                theme: "light1", // "light1", "light2", "dark1", "dark2"
                title:{
                    text: "'.$title.' '.$from.' - '.$to.'"
                },
                axisX:{
                    title : "Ilość transportowana"
                   },           
                   axisY:{
                    title : "Dni"
                   },
                data: [{
                    type: "'.$type.'", //change type to bar, line, area, pie, etc
                    indexLabel: "{x}", //Shows y value on all Data Points
                    indexLabelFontColor: "#5A5757",
                    indexLabelPlacement: "outside",
                    dataPoints: '.$data.'
                }]
            });
            chart.render();
            });
            ');
}

public function postal(){
    $form = & $this->init_module('Libs/QuickForm'); 
    $groups = Utils_CommonDataCommon::get_array("Companies_Groups");
    foreach($groups as $key => $val){
        $groups[$key] = __($val);
    }
    print("KORESPONDENCJA GRUPOWA <BR>");
    $form->addElement("select",'commonDataKey', 'Wybierz grupę do korespondencji: ',$groups );
    $form->addElement("submit", "submit", "Wygeneruj listę");
    $form->display();
    if($form->validate()){
        $valuesForm = $form->exportValues();
        $href = 'href="modules/driversRaport/korespondencja.php?'.http_build_query(array('cid'=>CID , 'commonDataKey' => $valuesForm['commonDataKey'])).'"';
        print("<a $href> Pobierz listę do korespondencji </a> <Br><br>");
    }
}

public function main(){
    // 
    
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
    $href = 'href="modules/driversRaport/excel.php?'.http_build_query(array('selected_month'=> $selected_month , 'year' => $selected_year , 'cid'=>CID)).'"';
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
        $href,
        null,
        3
    );
    $date_start = date("$selected_year-$selected_month-01");
    $date_end = date("$selected_year-".$selected_month."-t",strtotime($date_start));

    $redable_format = date("F",strtotime($date_start));
    print("<Br><Br> RAPORT KIEROWCÓW <BR>");
    print("Wybrany miesiąc - ".(__($redable_format))." $selected_year");
    }
}