<?php

namespace Fengxw\Excel;

class ExportExcel
{
    /**
     * object of phpExcel.
     */
    public $excelObj;

    public function __construct()
    {
        $this->excelObj = new \PHPExcel();
    }

    public static function getInstance()
    {
        return new self();
    }

    /**
     * Export excel simple.
     *
     * @param string $header
     * @param array  $title     ['column1', 'column2']
     * @param array  $data,     e.g. [['value1', 'value2']]
     *                          data is a two-dimension array, the order of column and value should be same.
     * @param string $sheetName
     *
     * @throws \PHPExcel_Exception
     */
    public function export(
        $header,
        $title,
        $data,
        $sheetName = 'sheet 1'
    ) {
        // set header
        $endOffset = self::numToLetter(count($title), true);
        $this->setHead($header, $endOffset.'1', $sheetName);

        // set title
        $this->setTitle($title);

        // set data
        $rowCount = 3;
        $this->setData($data, $rowCount);

        // set global style
        $rowCount = count($data) + 2;
        $endCell = $endOffset.$rowCount;
        $this->setDefaultStyle('A2', $endCell);

        // export
        $this->output(time().'.xls');
    }

    /**
     * Set excel header.
     *
     * @param $headTxt
     * @param $endCell
     * @param string $sheetName
     * @param string $startCell
     * @param int    $height
     * @param array  $excelStyle
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setHead(
        $headTxt,
        $endCell,
        $sheetName = 'sheet 1',
        $startCell = 'A1',
        $height = 40,
        $excelStyle = []
    ) {
        $this->excelObj->setActiveSheetIndex(0);
        $this->excelObj->getActiveSheet()->setTitle($sheetName);

        $styleHeader = [
            'start' => $startCell,
            'end' => $endCell,
            'height' => $height,
        ];

        if (!empty($excelStyle)) {
            $styleHeader['style'] = $excelStyle;
        } else {
            $styleHeader['style'] = [
                'borders' => [
                    'allborders' => [
                        'style' => \PHPExcel_Style_Border::BORDER_THIN,
                    ],
                ],
                'font' => [
                    'bold' => true,
                    'size' => 20,
                ],
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical' => \PHPExcel_Style_Alignment::VERTICAL_TOP,
                    'wrap' => true,
                ],
            ];
        }

        $this->setStyle($styleHeader)
            ->setCell($headTxt, $startCell, $endCell);

        return $this;
    }

    /**
     * Set excel columns.
     *
     * @param array $title
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setTitle($title)
    {
        $this->setData($title, 2);

        return $this;
    }

    /**
     * Set excel data.
     *
     * @param $data
     * @param $rowCount
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setData($data, $rowCount, $startColumn = 'A')
    {
        $this->excelObj
            ->getActiveSheet()
            ->fromArray(
                $data,
                null,
                $startColumn.$rowCount,
                true
            );

        return $this;
    }

    /**
     * Set cell.
     *
     * @param $startCell
     * @param $endCell
     * @param $value
     * @param $style
     *
     * @throws \PHPExcel_Exception
     * @return $this
     */
    public function setCell($value, $startCell, $endCell = null, $style = [])
    {
        $this->excelObj->getActiveSheet()->setCellValue($startCell, $value);

        if ($endCell) {
            $this->excelObj->getActiveSheet()->mergeCells($startCell.':'.$endCell);
        }

        if (!empty($style)) {
            $coordinate = $endCell ? $startCell.':'.$endCell : $startCell;

            $this->excelObj->getActiveSheet()->getStyle();
            $this->excelObj
                ->getActiveSheet()
                ->getStyle($coordinate)
                ->applyFromArray($style);
        }

        return $this;
    }

    /**
     * Set key cell.
     *
     * @param $key
     * @param $value
     * @param $startCell
     * @param string $endCell
     * @param string $mediumCell
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setKeyCell($key, $value, $startCell, $endCell, $mediumCell = '')
    {
        $this->setCell($key, $startCell, $mediumCell);

        if (!$mediumCell) {
            $mediumCell = $startCell;
        }

        $mediumCell = self::incrColumn($mediumCell);

        $this->setCell($value, $mediumCell, $endCell);

        return $this;
    }

    /**
     * Set style.
     *
     * @param $style
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setStyle($style)
    {
        $startCell = $style['start'];
        $endCell = $style['end'];
        $excelStyle = $style['style'];

        $this->excelObj->getActiveSheet()->getStyle();
        $this->excelObj
            ->getActiveSheet()
            ->getStyle($startCell.':'.$endCell)
            ->applyFromArray($excelStyle);

        // set width of column
        if ($style['width']) {
            $sColumn = substr($startCell, 0, 1);
            $eColumn = substr($endCell, 0, 1);

            foreach (range($sColumn, $eColumn) as $item) {
                $this->excelObj
                    ->getActiveSheet()
                    ->getColumnDimension($item)
                    ->setWidth($style['width']);
            }
        }

        //set height of row
        if ($style['height']) {
            $sRow = self::findNum($startCell);
            $eRow = self::findNum($endCell);

            foreach (range($sRow, $eRow) as $item) {
                $this->excelObj
                    ->getActiveSheet()
                    ->getRowDimension($item)
                    ->setRowHeight($style['height']);
            }
        }

        return $this;
    }

    /**
     * Set default style.
     *
     * @param $startCell
     * @param $endCell
     * @param int   $width
     * @param int   $height
     * @param array $excelStyle
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setDefaultStyle($startCell, $endCell, $width = 15, $height = 30, $excelStyle = [])
    {
        $style = [
            'start' => $startCell,
            'end' => $endCell,
            'width' => $width,
            'height' => $height,
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

        if (!empty($excelStyle)) {
            $style['style'] = $excelStyle;
        }

        $this->setStyle($style);

        return $this;
    }

    /**
     * Output excel file.
     *
     * @param $filename
     *
     * @throws \PHPExcel_Exception
     */
    public function output($filename)
    {
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$filename);
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($this->excelObj, 'Excel5');
        $objWriter->save('php://output');
    }

    /**
     * increase column of cell.
     *
     * @param $cell
     *
     * @return string
     */
    public static function incrColumn($cell)
    {
        $column = substr($cell, 0, 1);
        $row = self::findNum($cell);
        ++$column;

        return $column.$row;
    }

    /**
     * Get num of cell.
     *
     * @param string $str
     *
     * @return string
     */
    public static function findNum($str = '')
    {
        $str = trim($str);
        if (empty($str)) {
            return '';
        }
        $result = '';
        for ($i = 0; $i < strlen($str); ++$i) {
            if (is_numeric($str[$i])) {
                $result .= $str[$i];
            }
        }

        return $result;
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

    /**
     * Set width Column
     *
     * @param $column
     * @param $width
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setColumn($column, $width)
    {
        // set width of column
        $this->excelObj
            ->getActiveSheet()
            ->getColumnDimension($column)
            ->setWidth($width);

        return $this;
    }

    /**
     * Set height of row
     *
     * @param $row
     * @param $height
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setHeight($row, $height)
    {
        // set width of column
        $this->excelObj
            ->getActiveSheet()
            ->getRowDimension($row)
            ->setRowHeight($height);

        return $this;
    }

    /**
     * @param $sheetName
     * @param $index
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function activeSheet($sheetName, $index = 0)
    {
        $this->excelObj->setActiveSheetIndex($index);
        $this->excelObj->getActiveSheet()->setTitle($sheetName);

        return $this;
    }
}