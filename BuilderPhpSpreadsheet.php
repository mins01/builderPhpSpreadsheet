<?php
namespace mins01\BuilderPhpSpreadsheet;


/**
 * 2024-07-30: 자동링크 처리.
 */
class BuilderPhpSpreadsheet{
    public $spreadsheet = null;
    public $default_width = 20;

    static $defStyles = [
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
        $rIdx = 1; // row idx

        if(isset($header[0][0])){
            $partConf = ($conf['header']??[])+$this->defConf['header'];
            $rIdx = $this->setSheetDataPart($sheet,$rIdx,$header,$partConf);
        }
        if(isset($body[0][0])){
            $partConf = ($conf['body']??[])+$this->defConf['body'];
            $rIdx = $this->setSheetDataPart($sheet,$rIdx,$body,$partConf);
        }
        if(isset($footer[0][0])){
            $partConf = ($conf['footer']??[])+$this->defConf['footer'];
            $rIdx = $this->setSheetDataPart($sheet,$rIdx,$footer,$partConf);
        }
        $sheet->setSelectedCell('A1');
        return $sheet;
    }

    /**
     * 시트에 데이터 파트를 적용
     */
    public function setSheetDataPart($sheet,$rIdx,$partValues,$partConf){
        //-- 값 설정
        $d = & $partValues;
        $firstCellCoord = 'A'.$rIdx;
        $sheet->fromArray($d,null,$firstCellCoord);

        if($partConf['autolink']){
            foreach($d as $dRIdx=>$dRow){
                foreach($dRow as $dCIdx => $dCol){
                    if( filter_var($dCol, FILTER_VALIDATE_URL) ){ //URL 인 경우
                        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($dCIdx+1).($dRIdx+$rIdx);
                        $sheet->getCell($coord)->getHyperlink()->setUrl($dCol);
                        $sheet->getStyle($coord)->applyFromArray(static::$defStyles['fontForLink']);
                    }
                    if(filter_var($dCol, FILTER_VALIDATE_EMAIL)){ //EMAIL 인 경우
                        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($dCIdx+1).($dRIdx+$rIdx);
                        $sheet->getCell($coord)->getHyperlink()->setUrl('mailto:'.$dCol);
                        $sheet->getStyle($coord)->applyFromArray(static::$defStyles['fontForLink']);
                    }
                }
            }
        }
        
        // exit;
        $rIdx+= count($d);
        $lastColIndex = max(array_map(function($arr){return count($arr);},$d));
        $lastColAlpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($lastColIndex);
        $lastCellCoord = $lastColAlpha.($rIdx-1);
        //-- 파트의 스타일
        if(isset($partConf['style'])) $sheet->getStyle("{$firstCellCoord}:{$lastCellCoord}")->applyFromArray($partConf['style']);
        //-- 컬럼의 스타일
        if($partConf && isset($partConf['columns'])){
            foreach($partConf['columns'] as $k => $column ){
                if(!isset($column)) continue;
                $alpha = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($k+1);
                $colCoord = $alpha.($rIdx - count($d)).':'.$alpha.($rIdx - 1);
                if(isset($column['style'])){
                    $sheet->getStyle($colCoord)->applyFromArray($column['style']);
                }
            }
        }
        unset($d);
        return $rIdx;
    }
    
}