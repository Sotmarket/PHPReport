<?php
/**
 * Интерфейс декоратора
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 05.08.13
 * @version    $Id: $
 */
interface IDecorator{
    public function __construct(
        PHPExcel_Worksheet $excelSheet
    );
    public function decorate();
}