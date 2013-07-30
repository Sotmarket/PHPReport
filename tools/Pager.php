<?php
/**
 * 
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 26.07.13
 * @version    $Id: $
 */
class Pager{
    protected $excelSheet;


    protected $pageHeight;
    protected $rowBounds = array();
    const INCH_FACTOR = 25.4; // inches to mm

    public function __construct(PHPExcel_Worksheet $excelSheet, $pageHeight=NULL){
        $this
            ->setExcelSheet($excelSheet)
            ->setPageHeight($pageHeight)
        ;
    }

    protected function setPageHeight($pageHeight)
    {
        $this->pageHeight = $pageHeight;
        return $this;
    }

    protected function getPageHeight()
    {
        if (NULL == $this->pageHeight){
            $pageModel = PageModel::getByExcelSheet($this->getExcelSheet());
            $this->pageHeight = $pageModel->getHeight();
        }
        return $this->pageHeight;
    }

    public function setRowBounds($rowBounds)
    {
        $this->rowBounds = $rowBounds;
        return $this;
    }

    public function getRowBounds()
    {
        if (0 == count ($this->rowBounds)){
            $sheet = $this->getExcelSheet();
            $highestRow = $sheet->getHighestRow()+5;
            $page = 1;
            $printMargins = $sheet->getPageMargins();
            $footer_top  =($printMargins->getTop()+$printMargins->getBottom())*self::INCH_FACTOR;

            $sum = $footer_top;
            // $footer_top  =25;
            // сумма высот колонок заголовка, переносимого на каждой странице
            $header = $this->getHeaderRowHeight();
            //$dimensions = $sheet->getRowDimensions();

            $count_per_page = 0;
            for($i=1; $i<=$highestRow; $i++){
                $height = $this->getRowHeight($i);

                $sum+=$height;
                $count_per_page++;
                 echo($i . ":".$height."\n");
                if ($sum > (($this->getPageHeight()))){
                    $this->rowBounds[$page]=$i-1;
                    //echo ("page:" . $page . "count_per_page" . ($count_per_page-1)."\n");
                    $page++;
                    $sum=$height+$footer_top-$header;
                    $count_per_page = 0;

                }
            }
            // last page
            //echo("lastsum:".$sum."\n"); echo($this->getPageHeight());

            //echo("lastsum:".$sum."\n"); echo($this->getPageHeight()); die();
            $this->rowBounds[$page]=$i;


        }
        return $this->rowBounds;
    }

    protected function getHeaderRowHeight(){
        $sheet = $this->getExcelSheet();
        $rowsRepeat = $sheet->getPageSetup()->getRowsToRepeatAtTop();
        $result = 0;
        if (is_array($rowsRepeat)){

            for ($k=$rowsRepeat[0]; $k<=$rowsRepeat[1]; $k++){
                $result+= $sheet->getRowDimension($k)->getRowHeight();
            }

        }
        return $result;
    }
    public function getHeightOfRowRange($rowStart =1, $rowFinish=null){
        $sheet = $this->getExcelSheet();
        if (null == $rowFinish){
            $highestRow = $sheet->getHighestRow();
        }
        else {
            $highestRow = $rowFinish;
        }
        $sum = 0;
        for($i=$rowStart; $i<=$highestRow; $i++){
            $height = $this->getRowHeight($i);
            $sum+= $height;
        }
        return $sum;
    }
    protected function getRowHeight ($row){
        //$row = $row-1;
        $sheet = $this->getExcelSheet();
        $dimensions = $sheet->getRowDimensions();

        if (isset($dimensions[$row])){
            $height = $dimensions[$row]->getRowHeight( );
        }
        else{
            $height = $this->getDefaultRowHeight();
        }
        return $height;
    }
    public function getDefaultRowHeight (){
        //return 16.5;
        //echo(PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->getExcelSheet()->getDefaultStyle()->getFont()));
        //die();
        return PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->getExcelSheet()->getDefaultStyle()->getFont());
       // return 12.5;
    }
    public function getFullDocumentHeight(){
        return $this->getHeightOfRowRange(1, null);
    }
    public function getCountPages(){
        return count($this->getRowBounds());
    }
    public function getPageOfRow($row){
        //$row = $row-1;
        $rowBounds = $this->getRowBounds();
        $result = 1;
        foreach ($rowBounds as $page=>$finishRow){
            if ($row<=$finishRow){
                $result = $page;
                break;
            }
        }
        return $result;
    }



    protected function setExcelSheet( PHPExcel_Worksheet $excelSheet)
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