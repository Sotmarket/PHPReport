<?php
/**
 * 
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 29.07.13
 * @version    $Id: $
 */

class PageDecorator {
    protected $excelSheet;
    protected $dataFinishRow;
    protected $totalStrings = 1;
    public function __construct (PHPExcel_Worksheet $excelSheet, $dataFinishRow,$totalStrings=1){
        $this->setExcelSheet($excelSheet)
             ->setDataFinishRow($dataFinishRow)
             ->setTotalStrings($totalStrings)
        ;

    }

    protected function setTotalStrings($totalStrings)
    {
        $this->totalStrings = $totalStrings;
        return $this;
    }

    protected function getTotalStrings()
    {
        return $this->totalStrings;
    }

    public function decorate(){
        $pager = new Pager($this->getExcelSheet()); // A4
        $sheet = $this->getExcelSheet();
        $finishRow = $this->getDataFinishRow();
        $pageLastDataRow = $pager->getPageOfRow( $finishRow ); //147
        $countPages = $pager->getCountPages();
        $rowBounds = $pager->getRowBounds();
        // Размер футера документа
        $need = $pager->getHeightOfRowRange($finishRow+$this->getTotalStrings()+1, null);
        // если футер не помещается  - то надо вставлять пустые строки
        $rowHeight = $sheet->getRowDimension($finishRow)->getRowHeight();
        //$have =
        $have = $pager->getHeightOfRowRange($finishRow+$this->getTotalStrings()+1, $rowBounds[$pageLastDataRow]);

        if ($pageLastDataRow<$pager->getCountPages()){
            $rowHeight = $sheet->getRowDimension($finishRow)->getRowHeight();

            $countRows = round(($have)/$rowHeight)+2; //magick

            $sheet->insertNewRowBefore($finishRow+1,$countRows);

            for($k=$finishRow+1; $k<$finishRow+1+$countRows; $k++){
                $sheet->getCellByColumnAndRow(0, $k)->setValue("_");
                $sheet
                    ->getStyleByColumnAndRow(0,$k)
                    ->getFont()
                    ->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_WHITE));
                $sheet->mergeCellsByColumnAndRow(1,$k,3,$k);
            }
        }
        return $sheet;
    }
    protected function setDataFinishRow($dataFinishRow)
    {
        $this->dataFinishRow = $dataFinishRow;
        return $this;
    }

    protected function getDataFinishRow()
    {
        return $this->dataFinishRow;
    }

    protected function setExcelSheet($excelSheet)
    {
        $this->excelSheet = $excelSheet;
        return $this;
    }

    protected function getExcelSheet()
    {
        return $this->excelSheet;
    }

}