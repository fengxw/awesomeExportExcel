<?php

include __DIR__.'/../src/ExportExcel.php';
include __DIR__.'/../vendor/autoload.php';

use Fengxw\Excel\ExportExcel;

class Test
{
    public function simple()
    {
        // define the arguments.
        $header = 'It is an header';
        $title = ['column1', 'column2'];
        $data = [['value1', 'value2']];

        // export excel
        ExportExcel::getInstance()->export($header, $title, $data);
    }

    public function complex()
    {
        // define the arguments.
        $header = 'It is an header';
        $title = ['column1', 'column2', 'column3', 'column4', 'column5', 'column6'];

        $globalStyle = [
            'start' => 'A2',
            'end' => 'F8',
            'width' => 14,
            'height' => 30,
            'style' => [
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical' => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    'wrap' => true,
                ],
                'borders' => [
                    'allborders' => [
                        'style' => \PHPExcel_Style_Border::BORDER_THIN,
                    ],
                ],
            ],
        ];

        // export
        ExportExcel::getInstance()
            ->setStyle($globalStyle)
            ->setHead($header, 'F1', 'test sheet')
            ->setTitle($title)
            ->setCell('Key1', 'A3')
            ->setCell('it\'s value1', 'B3', 'C3')
            ->setCell('Key2', 'D3')
            ->setCell('it\'s value2', 'E3', 'F3')
            ->setKeyCell('Key3', 'it\'s value3', 'A4', 'F4')
            ->setKeyCell('Key4', 'it\'s value4', 'A5', 'F5', 'C5')
            ->setCell('Key5', 'A6', 'A8')
            ->setCell('it\'s value5', 'B6', 'F8')
            ->output(time().'.xls');
    }
}

$test = new Test();
$test->simple();
//$test->complex();
