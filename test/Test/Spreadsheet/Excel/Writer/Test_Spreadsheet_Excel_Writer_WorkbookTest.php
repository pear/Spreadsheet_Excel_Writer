<?php

/**
 * @author stev leibelt <artodeto@bazzline.net>
 * @since 2016-01-17
 */
class Test_Spreadsheet_Excel_Writer_WorkbookTest extends Test_Spreadsheet_Excel_WriterTestCase
{
    public function testSetVersion()
    {
        $workbook = $this->getNewWorkbook();

        $workbook->setVersion(8);
    }
}