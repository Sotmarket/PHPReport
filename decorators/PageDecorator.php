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
    protected $dataStartRow;
    protected $totalStrings = 1;
    public function __construct (
            PHPExcel_Worksheet $excelSheet,
            $dataStartRow,
            $dataFinishRow,

            $totalStrings=1
        ){
        $this->setExcelSheet($excelSheet)
             ->setDataFinishRow($dataFinishRow)
             ->setDataStartRow($dataStartRow)
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

    protected function setDataStartRow($dataStartRow)
    {
        $this->dataStartRow = $dataStartRow;
        return $this;
    }

    protected function getDataStartRow()
    {
        return $this->dataStartRow;
    }

    public function decorate(){
        $pager = new Pager($this->getExcelSheet());
        $sheet = $this->getExcelSheet();
        $a = $sheet->getRowDimensions();
        //print_r($sheet->getRowDimensions());
        $defaultFont = $sheet->getDefaultStyle()->getFont();
        for ($k=$this->getDataStartRow(); $k<=$this->getDataFinishRow();$k++){
            $a[$k]->setRowHeight($defaultFont->getSize());
        }
        $countPages = $pager->getCountPages();
        $rowBounds = $pager->getRowBounds();

        if (1 == $countPages){
            $rows = $sheet->getPageSetup()->getRowsToRepeatAtTop();
            $sheet->getPageSetup()->setRowsToRepeatAtTop(array(0=>0,0=>0));

            return $sheet;
        }

        $finishRow = $this->getDataFinishRow()+$this->getTotalStrings();
        $pageLastDataRow = $pager->getPageOfRow( $finishRow );

        $countRows = 0;
        if (
            ($rowBounds[$pageLastDataRow] == $finishRow)

        ){
            // если мы на последнем месте, нужно перенести

            $indexStart = $finishRow-$this->getTotalStrings();
            $countRows = 2;
            $have= 0;
            $rowHeight = 0;

        }elseif ($pageLastDataRow>1 && $rowBounds[$pageLastDataRow-1] == $finishRow-1){
            $indexStart = $finishRow-1;
            $countRows = 1;
            $have= 0;
            $rowHeight = 0;
        }
        elseif ($pageLastDataRow<$pager->getCountPages()){
            $indexStart = $finishRow-$this->getTotalStrings();
            $rowHeight = $sheet->getRowDimension($indexStart)->getRowHeight();
            //$have =

            $have = $pager->getHeightOfRowRange($finishRow+1, $rowBounds[$pageLastDataRow]);
            $countRows = round(($have)/$rowHeight)+2; //magick
        }
        //return $sheet;
        if ($countRows){

            $sheet->insertNewRowBefore($indexStart,$countRows);

            for($k=$indexStart; $k<$indexStart+$countRows; $k++){
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

    protected function setExcelSheet(PHPExcel_Worksheet $excelSheet)
    {
        $this->excelSheet = $excelSheet;
        return $this;
    }

    /**
     * @return PHPExcel_Worksheet
     */
    protected function getExcelSheet()
    {
        return $this->excelSheet;
    }

}