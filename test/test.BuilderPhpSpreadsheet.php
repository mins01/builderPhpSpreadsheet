<?
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

require_once('../vendor/autoload.php');
// require_once('../src/Builder.php');
use mins01\BuilderPhpSpreadsheet\BuilderPhpSpreadsheet;

$bps = new BuilderPhpSpreadsheet();

$conf = [
    'sheet'=>[
        'title'=>"기본설정",
        'cellSpans'=>[
            // [cellColumnIndex(1+),cellRowIndex(1+),colspan(0+),rowspan(0+)]
            [1,1,2,1],
        ]
    ],
    'body'=>[
        'cellSpans'=>[
            // [cellColumnIndex(1+),cellRowIndex(1+),colspan(0+),rowspan(0+)]
            [3,1,2,1],
        ],
        'mergeCells'=>[
            // [cellColumnIndex_1(1+),cellRowIndex_1(1+),cellColumnIndex_2(1+),cellRowIndex_2(1+)]
            [5,1,7,2],
        ]
    ]
    
];
$header = [
    ['해더1-'.time(),'해더2','해더3','해더4','해더5','해더6','해더7','해더8','해더9','해더10','해더11','해더12','해더13',],
    ['http://www.mins01.com/','test@mins01.com','해더:2-3','해더:2-4','해더:2-5','해더:2-6','해더:2-7','해더:2-8','해더:2-9','해더:2-10','해더:2-11','해더:2-12','해더:2-13',]
];
$footer = [
    ['풋터1','풋터2','풋터3','풋터4','풋터5','풋터6','풋터7','풋터8','풋터9','풋터10','풋터11','풋터12','풋터13',],
    ['http://www.mins01.com/','test@mins01.com','풋터:2-3','풋터:2-4','풋터:2-5','풋터:2-6','풋터:2-7','풋터:2-8','풋터:2-9','풋터:2-10','풋터:2-11','풋터:2-12','풋터:2-13',]
];
$body = [
    [1,2,3,4,5,6,7,8,9,10],
    ['http://www.mins01.com/','test@mins01.com','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
];
$bps->setSheetData(0,$conf,$body,$header,$footer);




$conf = [
    'sheet'=>[
        'title'=>"컬럼너비만 설정",
        'columns'=>[
            ['width'=>10],
            ['width'=>20],
            ['width'=>30],
            ['width'=>40],
            ['width'=>'auto'],
            ['width'=>30],
            ['width'=>20],
            ['width'=>10],
            ['width'=>5],
            [],
            [],
        ],
    ],
];
$header = [
    ['해더1-'.time(),'해더2','해더3','해더4','해더5','해더6','해더7','해더8','해더9','해더10','해더11','해더12','해더13',],
    ['해더:2-1','해더:2-2','해더:2-3','해더:2-4','해더:2-5','해더:2-6','해더:2-7','해더:2-8','해더:2-9','해더:2-10','해더:2-11','해더:2-12','해더:2-13',]
];
$footer = [
    ['풋터1','풋터2','풋터3','풋터4','풋터5','풋터6','풋터7','풋터8','풋터9','풋터10','풋터11','풋터12','풋터13',],
    ['풋터:2-1','풋터:2-2','풋터:2-3','풋터:2-4','풋터:2-5','풋터:2-6','풋터:2-7','풋터:2-8','풋터:2-9','풋터:2-10','풋터:2-11','풋터:2-12','풋터:2-13',]
];
$body = [
    ['10','20','30','40','auto','30','20','10','5','너비설정없음','너비설정없음','설정없음','설정없음',],
    ['데이터1','22222222','2024-07-15','설정이 없으면 기본너비 사용','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
    ['데이터1','22222222','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],

    
];
// downloadXlsx($fn,$conf,$body,$header,$footer);
$bps->setSheetData(1,$conf,$body,$header,$footer);



$conf = [
    'sheet'=>[
        'title'=>"복잡하게 설정",
        'columns'=>[
            ['width'=>10],
            ['width'=>20],
            ['width'=>30],
            ['width'=>40],
            ['width'=>'auto'],
            [],
            [],
            [],
            [],
            [],
            [],
        ],
    ],
    
    'header'=>[
        'columns'=>[
            ['style'=>[ 'font' => [ 'color'=>[ 'argb'=>PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED, ], ], ],],
        ]
    ],
    'footer'=>[
        'columns'=>[
            null,
            null,
            ['style'=>[ 'font' => [ 'color'=>[ 'argb'=>PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLUE, ], ], ],],
        ],
        'style'=>[ 
            'font' => [ 
                'bold'=>true, 
                'color'=>[
                    'argb'=>PhpOffice\PhpSpreadsheet\Style\Color::COLOR_GREEN,
                ]
            ], 
        ],

    ],
    'body'=>[
        'columns'=>[
            ['style'=>['numberFormat'=>['formatCode'=>PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT,]]],
            ['style'=>['numberFormat'=>['formatCode'=>PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER,]]],
            ['style'=>[
                    'numberFormat'=>['formatCode'=>PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDD,],
                    'alignment' => [
                        'horizontal' => PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                        'vertical' => PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                    ],
                ]
            ],
            null,
            ['style'=>[
                    'alignment' => [ 'horizontal' => PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT, ],
                ]
            ],
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
        ]
    ]
];
$header = [
    ['해더1-'.time(),'해더2','해더3','해더4','해더5','해더6','해더7','해더8','해더9','해더10','해더11','해더12','해더13',],
    ['해더:2-1','해더:2-2','해더:2-3','해더:2-4','해더:2-5','해더:2-6','해더:2-7','해더:2-8','해더:2-9','해더:2-10','해더:2-11','해더:2-12','해더:2-13',],
    [null,null,null,'해더컬럼스타일'],
];
$footer = [
    ['풋터1','풋터2','풋터3','풋터4','풋터5','풋터6','풋터7','풋터8','풋터9','풋터10','풋터11','풋터12','풋터13',],
    ['풋터:2-1','풋터:2-2','풋터:2-3','풋터:2-4','풋터:2-5','풋터:2-6','풋터:2-7','풋터:2-8','풋터:2-9','풋터:2-10','풋터:2-11','풋터:2-12','풋터:2-13',],
    [null,null,null,'풋터컬럼스타일,풋터전체스타일재설정'],
];
$body = [
    ['데이터1','22222222','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
    ['데이터1','22222222','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
    ['데이터1','22222222','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
    [null,null,null,'컬럼별 formatCode 설정'],

    
];
$bps->setSheetData(2,$conf,$body,$header,$footer);



//-- 셀 머지 예제. (BuilderPhpSpreadsheet 에서는 지원안하니깐.. 수동으로 처리한다.)
$spreadsheet = $bps->spreadsheet;
$sheet = $spreadsheet->getSheet(2);
$sheet->mergeCells('A4:C6');
$sheet->setCellValue('A4','셀 머지');
$sheet->getStyle('A4')->applyFromArray(BuilderPhpSpreadsheet::$defStyles['alignmentCenterCenter']);




$conf = [
    'sheet'=>[
        'title'=>"오토링크 설정",
    ],
    'body'=>[
        'autolink'=>true //기본 true
    ],
    'header'=>[
        'autolink'=>true //기본 false
    ],
    'footer'=>[
        'autolink'=>true //기본 false
    ],

];
$header = [
    ['해더1-'.time(),'해더2','해더3','해더4','해더5','해더6','해더7','해더8','해더9','해더10','해더11','해더12','해더13',],
    ['http://www.mins01.com/','test@mins01.com','해더:2-3','해더:2-4','해더:2-5','해더:2-6','해더:2-7','해더:2-8','해더:2-9','해더:2-10','해더:2-11','해더:2-12','해더:2-13',]
];
$footer = [
    ['풋터1','풋터2','풋터3','풋터4','풋터5','풋터6','풋터7','풋터8','풋터9','풋터10','풋터11','풋터12','풋터13',],
    ['http://www.mins01.com/','test@mins01.com','풋터:2-3','풋터:2-4','풋터:2-5','풋터:2-6','풋터:2-7','풋터:2-8','풋터:2-9','풋터:2-10','풋터:2-11','풋터:2-12','풋터:2-13',]
];
$body = [
    [1,2,3,4,5],
    ['http://www.mins01.com/','test@mins01.com','2024-07-15','데이터4','오른쪽 정렬','데이터6','데이터7','데이터8','데이터9','데이터10','데이터11','데이터12','데이터13',],
];
$bps->setSheetData(3,$conf,$body,$header,$footer);



//--- xlsx 이면 다운로드, html 이면 브라우저에서 그냥 보여진다.
// $bps->download('testExcel','html');
// $bps->download('testExcel','xlsx');
$bps->save('testExcel','xlsx');
$bps->save('testExcel','html');

