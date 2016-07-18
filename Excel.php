<?php

use PHPExcel;
use PHPExcel_Style_Alignment;
use PHPExcel_IOFactory;

class Excel {

    //能读写的文件格式类型代号
    const EXCEL5 = 'Excel5';
    const EXCEL2007 = 'Excel2007';
    const EXCEL2003XML = 'Excel2003XML';
    const OOCALC = 'OOCalc';
    const SYLK = 'SYLK';
    const GNUMERIC = 'Gnumeric';
    const HTML = 'HTML';
    const CSV = 'CSV';
    //样式中数据位置常量
    const CENTER = 'center';
    const LEFT = 'left';
    const RIGHT = 'right';
    const TOP = 'top';
    const BOTTOM = 'bottom';

    //PHPExcel对象
    public static $objPHPExcel;
    //要输出的文件名称
    private static $filename;

    public function __construct() {
        self::$objPHPExcel = new PHPExcel();
    }

    /**
     * excel文件导出功能封装，依赖于Excel扩展包。详情请参考vendor/phpoffice/phpexcel/Classes/PHPExcel.php
     * 
     * @param array $data 需要导出的数据，格式是一个二维数组，类似于：
     * array (
      0 =>
      array (
      0 => 'zhangsan',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      1 =>
      array (
      0 => '总计',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      )
     * 另外，此参数也提供了headerList参数，它在$headerList参数为空的情况下可以覆盖$headerList类似于：
     * array (
      0 =>
      array (
      0 => 'zhangsan',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      1 =>
      array (
      0 => '总计',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      'headerList' =>
      array (
      0 => '交易人员',
      1 => '稽查结果统计（真/待确认）',
      2 => '稽查结果统计（假）',
      3 => '占比（真/待确认）',
      4 => '占比（假）',
      ),
      )
     * 注意，在不需要设置header（表单头）样式的情况下，完全可以将所有数据存入data而不需要设置headerList，
     * data中不要有headerList为key的数据，$headerList也保持为空即可。
     * @param array $headerList header表单样式头数据，如果不需要单独设置表单头样式，可以默认为空。样式参考上面示例。
     * @param string $title 表单title，也即是sheet的标题。
     * @param string $filename 下载下来的时候文件的名称，
     * 可通过setFileName()方法设置默认名称或者不传/为空都会自动调用setFileName()获取默认filename;
     * @param boolean $output 是否直接输出下载。true：直接下载；false：不直接下载。默认为true。
     * @param string $excelType 下载的文件的格式。有Excel5，Excel2007等。
     * @param array $options 提供额外参数，对样式等进行修改的可能。
     * 目前只实现了对header的样式的调整，header样式调整参数示例：
     * ['header_style' => ['hstyle' => 'center', 'vstyle' => 'center']]
     * @return type
     */
    public static function export($data, $headerList = [], $title = '', $filename = '', $output = true, $excelType = 'Excel5', $options = []) {
        self::$objPHPExcel = self::$objPHPExcel? : (new PHPExcel());
        //设置当前活动的sheet
        self::$objPHPExcel->setActiveSheetIndex(0);

        //设置sheet名字
        if ($title) {
            self::$objPHPExcel->getActiveSheet()->setTitle($title);
        }
        //设置默认行高
        self::$objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);

        $headerList = self::getHeaderList($data, $headerList);

        $startRow = 1;
        if ($headerList) {
            if (isset($options['header_style'])) {
                self::setHeader($headerList, self::$objPHPExcel, $options['header_style']);
            } else {
                self::setHeader($headerList, self::$objPHPExcel);
            }
            $startRow = 2;
        }

        self::setCellValue($data, self::$objPHPExcel, $startRow);

        $filename = $filename? : self::getFileName();

        $objWriter = PHPExcel_IOFactory::createWriter(self::$objPHPExcel, $excelType);
        if ($output) {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename=' . $filename . ' ');
            header('Cache-Control: max-age=0');
            $objWriter->save('php://output');
        } else {
            return $objWriter;
        }
    }

    /**
     * 设置文件名称
     * 默认情况下以时间戳加5位随机数拼接为文件名
     * 
     * @param string $filename 需要设置的文件名称
     * @return \app\components\Excel
     */
    public static function setFileName($filename = '') {
        self::$filename = $filename? : date("YmdHis") . rand(10000, 99999) . ".xls";
    }

    /**
     * 获取文件名filename，如果filename为空，则尝试调用一次setFileName再返回。
     * 
     * @return type
     */
    public static function getFileName() {
        if (!self::$filename) {
            self::setFileName();
        }
        return self::$filename;
    }

    /**
     * 设置每个cell的值
     * 
     * @param array $data 纯净的能直接设置到cell中的二维数组。类似于：
     * array (
      0 =>
      array (
      0 => 'zhangsan',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      1 =>
      array (
      0 => '总计',
      1 => 2,
      2 => 2,
      3 => '50%',
      4 => '50%',
      ),
      )
     * @param PHPExcel self::$objPHPExcel 需要操作的PHPExcel对象
     * @param int $startRow 提供可以调整数据从哪一行开始循环写入的开始行数值。默认从第一行（1）开始。
     * @param type $startColumn 提供可以调整数据从哪一列开始循环写入的开始列数值。默认从第一列（0）开始。
     * @return PHPExcel
     */
    public static function setCellValue($data, PHPExcel $objPHPExcel = null, $startRow = 1, $startColumn = 0) {
        self::$objPHPExcel = $objPHPExcel? : self::$objPHPExcel;
        foreach ($data as $rowKey => $row) {
            $rowIndex = $startRow ? $rowKey + $startRow : $rowKey;
            foreach ($row as $columnKey => $column) {
                $columnIndex = $startColumn ? $columnKey + $startColumn : $columnKey;
                self::$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($columnIndex, $rowIndex, $column);
                $columnIndex++;
            }
            $rowIndex++;
        }
        return self::$objPHPExcel;
    }

    /**
     * 对data中的headerList和$headerList数据进行调整。
     * $headerList为空的时候data中的headerList有值则赋值给$headerList.
     * 最终删除data中的headerList数据。保证data中数据的纯净性。 
     * 
     * @param array $data 数据示例参照export函数
     * @param array $headerList 数据示例参照export函数
     * @return type
     */
    public static function getHeaderList(&$data, $headerList = []) {
        if (isset($data['headerList']) && !empty($data['headerList'])) {
            if (empty($headerList)) {
                $headerList = $data['headerList'];
            }
            unset($data['headerList']);
        }
        return $headerList;
    }

    /**
     * 设置header数据
     * 
     * 
     * @todo 现只提供最多26个列的样式支持。现只对横排的列进行了支持，尚未支持竖排或者既有竖排又有横排的。
     * @param array $headerList 数据示例参照export函数
     * @param PHPExcel $objPHPExcel
     * @param type $options
     */
    public static function setHeader($headerList, PHPExcel $objPHPExcel, $options = []) {
        self::$objPHPExcel = $objPHPExcel? : self::$objPHPExcel;
        $letterArr = range('A', 'Z');
        foreach ($headerList as $headerIndex => $header) {
            $objStyle = self::$objPHPExcel
                    ->getActiveSheet()
                    ->setCellValueByColumnAndRow($headerIndex, 1, $header)
                    ->getStyle($letterArr[$headerIndex] . '1')
                    ->getAlignment();
            if (isset($options['hstyle']) && $options['hstyle']) {
                self::setHorizontal($objStyle, $options['hstyle']);
            }
            if (isset($options['vstyle']) && $options['vstyle']) {
                self::setVertical($objStyle, $options['vstyle']);
            }
        }
    }

    /**
     * 设置单元格（cell）的横排样式，有左右中三种值
     * 
     * @param PHPExcel_Style_Alignment $objStyle
     * @param string $style center、left、right
     * @return PHPExcel_Style_Alignment
     */
    public static function setHorizontal(PHPExcel_Style_Alignment $objStyle, $style = 'center') {
        switch ($style) {
            case self::LEFT:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                break;
            case self::CENTER:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                break;
            case self::RIGHT:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                break;
            default :
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                break;
        }
        return $objStyle;
    }

    /**
     * 设置单元格（cell）的竖排样式，有上中下三种值
     * 
     * @param PHPExcel_Style_Alignment $objStyle
     * @param string $style center、top、bottom
     * @return PHPExcel_Style_Alignment
     */
    public static function setVertical(PHPExcel_Style_Alignment $objStyle, $style = 'center') {
        switch ($style) {
            case self::BOTTOM:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_BOTTOM);
                break;
            case self::CENTER:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                break;
            case self::TOP:
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_TOP);
                break;
            default :
                $objStyle->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                break;
        }
        return $objStyle;
    }

}
