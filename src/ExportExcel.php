<?php

namespace Fengxw\ExportExcel;

class ExportExcel
{
    static function getInstance()
    {
        return new self();
    }

    /**
     * export excel.
     *
     * @param string $header
     * @param array $title ['column1', 'column2']
     * @param array $data, e.g. [['value1', 'value2']]
     *   data is a two-dimension array, the order of column and value should be same.
     * @param string $sheetName
     */
    public function export(
        $header,
        $title,
        $data,
        $sheetName = 'sheet 1'
    ) {
        $objPHPExcel = new \PHPExcel();

        // set title
        $endOffset = self::numToLetter(count($title), true);
        $this->setExcelTitle($objPHPExcel, $title, $endOffset);

        // set header
        $this->setExcelHead($objPHPExcel, $header, $endOffset.'1', $sheetName);

        // set data
        $rowCount = 3;
        $this->setExcelData($objPHPExcel, $data, $rowCount);

        // set global style
        $rowCount = count($data) + 2;
        $endCell = $endOffset.$rowCount;
        $this->setExcelGlobalStyle($objPHPExcel, 'A2', $endCell);

        // export
        $filename = time().'.xls';
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$filename);
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
    }

    /**
     * Set excel header.
     *
     * @param \PHPExcel $objPHPExcel
     * @param $HeadTxt
     * @param $endCell
     * @param $sheetName
     */
    public function setExcelHead($objPHPExcel, $HeadTxt, $endCell, $sheetName = 'sheet 1')
    {
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setTitle($sheetName);

        $startCell = 'A1';
        $objPHPExcel->getActiveSheet()->setCellValue($startCell, $HeadTxt);
        $objPHPExcel->getActiveSheet()->mergeCells($startCell.':'.$endCell);

        $styleHeader = [
            'borders' => [
                'allborders' => [
                    'style' => \PHPExcel_Style_Border::BORDER_THIN,
                ],
            ],
            'font' => [
                'bold' => true,
                'size' => 14,
            ],
            'alignment' => [
                'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                'vertical' => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
                'wrap' => true,
            ],
        ];

        $objPHPExcel
            ->getActiveSheet()
            ->getStyle($startCell.':'.$endCell)
            ->applyFromArray($styleHeader);

        //set hight of row as 40
        $objPHPExcel
            ->getActiveSheet()
            ->getRowDimension('1')
            ->setRowHeight(40);
    }

    /**
     * Set excel columns.
     *
     * @param \PHPExcel $objPHPExcel
     * @param array $title
     * @param $endOffset
     */
    public function setExcelTitle($objPHPExcel, $title, $endOffset)
    {
        $objPHPExcel->getActiveSheet()->fromArray($title, null, 'A2', true);

        // set width of column as 15
        foreach (range('A', $endOffset) as $item) {
            $objPHPExcel->getActiveSheet()->getColumnDimension($item)->setWidth(15);
        }
    }

    /**
     * Set excel data.
     *
     * @param \PHPExcel $objPHPExcel
     * @param $data
     * @param $rowCount
     */
    public function setExcelData($objPHPExcel, $data, $rowCount)
    {
        $objPHPExcel
            ->getActiveSheet()
            ->fromArray(
                $data,
                null,
                'A'.$rowCount,
                true
            );
    }

    /**
     * Set global style.
     *
     * @param \PHPExcel $objPHPExcel
     * @param $startCell
     * @param $endCell
     */
    public function setExcelGlobalStyle($objPHPExcel, $startCell, $endCell)
    {
        $objPHPExcel->getActiveSheet()->getStyle();
        $styleTitle = [
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
        ];

        $objPHPExcel
            ->getActiveSheet()
            ->getStyle($startCell.':'.$endCell)
            ->applyFromArray($styleTitle);
    }

    /**
     * convert num to letter
     * e.g: 1->a, 26->z, 27->aa.
     *
     * @param $num // minimum 1
     * @param bool $uppercase
     *
     * @return string
     */
    public static function numToLetter($num, $uppercase = false)
    {
        $num -= 1;
        for ($letter = ''; $num >= 0; $num = intval($num / 26) - 1) {
            $letter = chr($num % 26 + 0x41).$letter;
        }

        return $uppercase ? strtoupper($letter) : $letter;
    }
}
