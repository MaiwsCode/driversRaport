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
}
