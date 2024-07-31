<?php
namespace mins01\BuilderPhpSpreadsheet;

/**
 * 2024-07-30: 자동링크 처리.
 * 2024-07-31: conf cellSpans 처리 추가.
 */

/**
 * BuilderPhpSpreadsheet
 */
class BuilderPhpSpreadsheet{    
    /**
     * spreadsheet
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet||NULL
     */
    public $spreadsheet = null;
    
    /**
     * 기본 너비
     *
     * @var int
     */
    public $default_width = 20;
    
    /**
     * 기본 스타일
     *
     * @var array
     */
    static public $defStyles = [
        'alignmentCenterCenter' => [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ],
        'fillForHeader' =>[
            'fill'=>[
                'fillType'=>\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => array('rgb' => 'C6EFCE'),
            ]
        ],
        'fillForFooter' =>[
            'fill'=>[
                'fillType'=>\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => array('rgb' => 'eeeeee'),
            ]
        ],
        'bordersAll'=>[
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '00000000'],
                ],
            ],
        ],
        'fontForLink'=>[
            'font'  => [
                    'color' => ['rgb' => '0000FF'],
                    'underline' => 'single',
            ],
        ]
    ];
    
    public $defConf = [];


    public function __construct() {
        $this->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $this->init();
    }

    private function init(){
        $this->spreadsheet->getProperties()->setCreator('BuilderPhpSpreadsheet')
        ->setLastModifiedBy('BuilderPhpSpreadsheet')
        ->setTitle('BuilderPhpSpreadsheet')
        ->setSubject('BuilderPhpSpreadsheet')
        ->setDescription('BuilderPhpSpreadsheet')
        ->setKeywords('BuilderPhpSpreadsheet')
        ->setCategory('BuilderPhpSpreadsheet')
        ->setDescription(
            "CREATED BY BuilderPhpSpreadsheet"
        );
        // $this->spreadsheet->removeSheetByIndex(1);

        $this->defConf = [
            'header'=>['autolink'=>false,'style'=>static::$defStyles['fillForHeader']+static::$defStyles['bordersAll']+static::$defStyles['alignmentCenterCenter']],
            'body'=>['autolink'=>true,'style'=>static::$defStyles['bordersAll'],],
            'footer'=>['autolink'=>false,'style'=>static::$defStyles['fillForFooter']+static::$defStyles['bordersAll']],
        ];
    }


    /**
     * 웹 다운로드
     *
     * @param mixed $filename
     * @param mixed $type='xlsx'
     * 
     * @return null
     * 
     */
    public function download($filename,$type='xlsx'){
        if($type=='xlsx'){
            // 다운로드 - xlsx
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
        }else if($type=='html'){
            // 테스트 출력부 - HTML
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($this->spreadsheet);    
            $writer->save('php://output');
        }
    }


    /**
     * 서버에 파일 저장
     *
     * @param mixed $filename
     * @param mixed $type='xlsx'
     * 
     * @return null
     * 
     */
    public function save($filename,$type='xlsx'){
        if($type=='xlsx'){
            // 다운로드 - xlsx
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
            $writer->save($filename.'.xlsx');
        }else if($type=='html'){
            // 테스트 출력부 - HTML
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($this->spreadsheet);    
            $writer->save($filename.'.html');
        } 
    }

    /**
     * [Description for setSheetDataFromSheet]
     *
     * @param int $idx 시트번호(0+)
     * @param mixed $conf=null 설정
     * @param mixed $body=null body 데이터
     * @param mixed $header=null header 데이터
     * @param mixed $footer=null footer 데이터
     * 
     * @return PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet 
     * 
     */
    public function setSheetData($idx,$conf=null,$body=null,$header=null,$footer=null){
        $sheet =  null;
        if($idx <= $this->spreadsheet->getSheetCount()-1){
            $sheet = $this->spreadsheet->getSheet($idx);
        }else{
            $sheet = $this->spreadsheet->createSheet($idx);
        }
        if(!$sheet){
            throw new \Exception( "Not exists sheet.({$idx})");
        }
        return $this->setSheetDataFromSheet($sheet,$conf,$body,$header,$footer);
    }

    /**
     * [Description for setSheetDataFromSheet]
     *
     * @param PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet 
     * @param mixed $conf=null 설정
     * @param mixed $body=null body 데이터
     * @param mixed $header=null header 데이터
     * @param mixed $footer=null footer 데이터
     * 
     * @return PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet 
     * 
     */
    public function setSheetDataFromSheet($sheet,$conf=null,$body=null,$header=null,$footer=null){
        if(!isset($this->default_width)) $sheet->getDefaultColumnDimension()->setWidth($this->default_width); // 각 시트에 기본 너비 설정함
        
        
        //-- 시트 설정
        if(isset($conf['sheet']['title'])){ $sheet->setTitle($conf['sheet']['title']);}
        //--- 시트 column 설정
        $cf = $conf['sheet']??null;
        if(isset($cf['columns'])){
            foreach($cf['columns'] as $k => $column){
                $cIdx = $k+1; //column idx
                if(isset($column['width'])){
                    if($column['width'] =='auto'){
                        $sheet->getColumnDimensionByColumn($cIdx)->setAutoSize(true);
                    }else{
                        $sheet->getColumnDimensionByColumn($cIdx)->setWidth($column['width']);
                    }
                }            
            }
        }
        //--- 시트 cellSpans
        if(isset($cf['cellSpans'])){
            $this->setSheetCellSpans($sheet,1,1,$cf['cellSpans']);
        }
        //--- 데이터 설정
        $cellRowIndex = 1; // cell row idx
        $cellColumnIndex = 1; // cell column idx

        if(isset($header[0][0])){
            $partConf = ($conf['header']??[])+$this->defConf['header'];
            $r = $this->setSheetDataPart($sheet,$cellColumnIndex,$cellRowIndex,$header,$partConf);
            $cellRowIndex = $r[0];
        }
        if(isset($body[0][0])){
            $partConf = ($conf['body']??[])+$this->defConf['body'];
            $r = $this->setSheetDataPart($sheet,$cellColumnIndex,$cellRowIndex,$body,$partConf);
            $cellRowIndex = $r[0];
        }
        if(isset($footer[0][0])){
            $partConf = ($conf['footer']??[])+$this->defConf['footer'];
            $r = $this->setSheetDataPart($sheet,$cellColumnIndex,$cellRowIndex,$footer,$partConf);
            $cellRowIndex = $r[0];
        }
        $sheet->setSelectedCell('A1');
        return $sheet;
    }

    /**
     * 시트에 데이터 파트를 적용
     *
     * @param PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet 
     * @param int $fromCellColumnIndex 시작 column (1+)
     * @param int $fromCellRowIndex 시작 row (1+)
     * @param array $partValues 데이터
     * @param array $partConf 설정
     * 
     * @return array [$nextCellRowIndex,$nexCellColumnIndex] [1+,1+]
     * 
     */
    public function setSheetDataPart($sheet,$fromCellColumnIndex,$fromCellRowIndex,$partValues,$partConf){
        //-- 값 설정
        $d = & $partValues;
        $nextCellRowIndex = $fromCellRowIndex + count($d);
        $firstCellCoord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($fromCellColumnIndex).$fromCellRowIndex;
        $sheet->fromArray($d,null,$firstCellCoord);

        if($partConf['autolink']){
            foreach($d as $dRI=>$dRow){
                foreach($dRow as $dCI => $dVal){
                    if( filter_var($dVal, FILTER_VALIDATE_URL) ){ //URL 인 경우
                        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($dCI+$fromCellColumnIndex).($dRI+$fromCellRowIndex);
                        $sheet->getCell($coord)->getHyperlink()->setUrl($dVal);
                        $sheet->getStyle($coord)->applyFromArray(static::$defStyles['fontForLink']);
                    }
                    if(filter_var($dVal, FILTER_VALIDATE_EMAIL)){ //EMAIL 인 경우
                        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($dCI+$fromCellColumnIndex).($dRI+$fromCellRowIndex);
                        $sheet->getCell($coord)->getHyperlink()->setUrl('mailto:'.$dVal);
                        $sheet->getStyle($coord)->applyFromArray(static::$defStyles['fontForLink']);
                    }
                }
            }
        }
        
        // exit;
        $lastColIndex = $fromCellColumnIndex+max(array_map(function($arr){return count($arr);},$d))-1;
        $lastColAlpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($lastColIndex);
        $lastCellCoord = $lastColAlpha.($nextCellRowIndex-1);
        $nexCellColumnIndex = $lastColIndex+1;
        //-- 파트의 스타일
        if(isset($partConf['style'])) $sheet->getStyle("{$firstCellCoord}:{$lastCellCoord}")->applyFromArray($partConf['style']);
        //-- 컬럼의 스타일
        if($partConf && isset($partConf['columns'])){
            foreach($partConf['columns'] as $k => $column ){
                if(!isset($column)) continue;
                $alpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($k+1);
                $colCoord = $alpha.($nextCellRowIndex - count($d)).':'.$alpha.($nextCellRowIndex - 1);
                if(isset($column['style'])){
                    $sheet->getStyle($colCoord)->applyFromArray($column['style']);
                }
            }
        }
        //---  cellSpans
        if(isset($partConf['cellSpans'])){
            $this->setSheetCellSpans($sheet,$fromCellColumnIndex,$fromCellRowIndex,$partConf['cellSpans']);
        }
        //---  mergeCells
        if(isset($partConf['mergeCells'])){
            $this->setSheetMergeCells($sheet,$fromCellColumnIndex,$fromCellRowIndex,$partConf['mergeCells']);
        }
        return [$nextCellRowIndex,$nexCellColumnIndex];
    }

    /**
     * row,column 기준 오른쪽, 아래로 합친다.
     *
     * @param mixed $sheet
     * @param mixed $fromCellColumnIndex 1+
     * @param mixed $fromCellRowIndex 1+
     * @param mixed $cellSpans [[cellColumnIndex(1+),cellRowIndex(1+),colspan(0+),rowspan(0+)]]
     * 
     * @return [type]
     * 
     */
    public function setSheetCellSpans($sheet,$fromCellColumnIndex,$fromCellRowIndex,$cellSpans){
        $mergeCells = [];
        foreach($cellSpans as $cellSpan){
            if($cellSpan[2] > 1 || $cellSpan[3]>1 ){
                $mergeCells[] = [$cellSpan[0],$cellSpan[1],$cellSpan[0]+$cellSpan[2]-1,$cellSpan[1]+$cellSpan[3]-1];
            }
        }
        if(count($mergeCells)){
            $this->setSheetMergeCells($sheet,$fromCellColumnIndex,$fromCellRowIndex,$mergeCells);
        }
    }

    /**
     * row1,column1 에서 row2,column2 방향으로 합친다.
     *
     * @param mixed $sheet
     * @param mixed $fromCellRowIndex
     * @param mixed $fromCellColumnIndex
     * @param mixed $mergeCells [[cellColumnIndex_1(1+),cellRowIndex_1(1+),cellColumnIndex_2(1+),cellRowIndex_2(1+)]]
     * 
     * @return [type]
     * 
     */
    public function setSheetMergeCells($sheet,$fromCellColumnIndex,$fromCellRowIndex,$mergeCells){
        foreach($mergeCells as $mergeCell){
            if($mergeCell[2] > 1 || $mergeCell[3]>1 ){
                $coord1 = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($fromCellColumnIndex+$mergeCell[0]-1).($fromCellRowIndex+$mergeCell[1]-1);
                $coord2 = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($fromCellColumnIndex+$mergeCell[2]-1).($fromCellRowIndex+$mergeCell[3]-1);
                $coords = $coord1.':'.$coord2;
                // print_r($coords);
                $sheet->mergeCells($coords);
            }
        }
    }
    
    
}