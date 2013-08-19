<?php
/**
 * ��������� � �������� �������� ��������� �������� �� $wordsByLine ����
 * � ������������� ������ �������
 * ����� ��� ������������� ���������
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 06.08.13
 * @version    $Id: $
 */

class CellHeightDecorator implements IDecorator{
    protected $startDataRow;
    protected $countData;
    protected $excelSheet;
    protected $wordsByLine;
    protected $dynamicColumn;
    public function __construct ( PHPExcel_Worksheet $excelSheet ){
        $this->setExcelSheet($excelSheet);
    }

    /**
     * ������������ excelSheet
     */
    public function decorate(){
        $sheet = $this->getExcelSheet();
        //print_r($sheet); die();
        $rowDimension = $sheet->getRowDimensions();
        $defaultFont = $sheet->getDefaultStyle()->getFont();
        $defSize = PHPExcel_Shared_Font::getDefaultRowHeightByFont($defaultFont);
        $finish  = $this->getStartDataRow()+$this->getCountData()-1;

        for ($t=$this->getStartDataRow(); $t<=$finish;$t++){
            $cell = $sheet->getCellByColumnAndRow($this->getDynamicColumn(),$t);
            $value = $cell->getValue();
            //$length = mb_strlen($value);
           // $chunks = round($length/40);
            $value = preg_replace('/(\v|\s)+/', ' ', $value);
            $nValue = wordwrap($value, $this->getWordsByLine(),"\n");
            $count  = substr_count ($nValue, "\n");
            $rowSize = $defSize*($count+1);

            $cell->setValue($nValue);
            if (!isset($rowDimension[$t])){

                $objDimension = new  PHPExcel_Worksheet_RowDimension($t);
                $objDimension->setRowHeight($rowSize);
                $rowDimension[$t] = $objDimension;
            }
            else {
                $rowDimension[$t]->setRowHeight($rowSize);
            }
        }
        $sheet->setRowDimensions($rowDimension);
    }

    /**
     * @param $dynamicColumn
     * @return $this
     */
    public function setDynamicColumn($dynamicColumn)
    {
        $this->dynamicColumn = $dynamicColumn;
        return $this;
    }

    protected function getDynamicColumn()
    {
        return $this->dynamicColumn;
    }

    /**
     * ���������� ���� � ������
     * @param $wordsByLine
     * @return $this
     */
    public function setWordsByLine($wordsByLine)
    {
        $this->wordsByLine = $wordsByLine;
        return $this;
    }

    protected function getWordsByLine()
    {
        return $this->wordsByLine;
    }

    /**
     * ���������� ����� ��� ���������
     * @param $countData
     * @return $this
     */
    public function setCountData($countData)
    {
        $this->countData = $countData;
        return $this;
    }

    protected function getCountData()
    {
        return $this->countData;
    }

    /**
     * @param $excelSheet
     * @return $this
     */
    protected function setExcelSheet($excelSheet)
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

    /**
     * ������ ������
     * @param $startDataRow
     * @return $this
     */
    public function setStartDataRow($startDataRow)
    {
        $this->startDataRow = $startDataRow;
        return $this;
    }

    protected function getStartDataRow()
    {
        return $this->startDataRow;
    }

}