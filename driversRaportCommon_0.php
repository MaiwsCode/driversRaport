<?php
defined("_VALID_ACCESS") || die('Direct access forbidden');
/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

class driversRaportCommon extends ModuleCommon {

    public static function menu() {
		return array(_M('Reports') => array('__submenu__' => 1, __('Miesięczny raport kierowców') => array(
	    'view'
			)));
	}
    public static function getCell($type){
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

}

