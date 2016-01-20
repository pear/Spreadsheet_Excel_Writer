<?php

/**
 * @author stev leibelt <artodeto@bazzline.net>
 * @since 2016-01-17
 */
class Test_Spreadsheet_Excel_WriterTestCase extends PHPUnit_Framework_TestCase
{
    /**
     * @param string $fileName
     * @return Spreadsheet_Excel_Writer_Workbook
     */
    protected function getNewWorkbook($fileName = 'my_workbook')
    {
        return new Spreadsheet_Excel_Writer_Workbook($fileName);
    }
}