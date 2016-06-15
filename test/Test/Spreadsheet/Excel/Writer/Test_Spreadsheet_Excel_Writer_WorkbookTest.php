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

        $before = get_object_vars($workbook);

        $workbook->setVersion(1);

        $this->assertEquals($before, get_object_vars($workbook), "Version 1 should not change internal state");

        $workbook->setVersion(8);

        $this->assertNotEquals($before, get_object_vars($workbook), "Version 8 should change internal state");

        return $workbook;
    }

    /**
     * @depends testSetVersion
     */
    public function testWriteSingleCell(Spreadsheet_Excel_Writer $workbook)
    {
        $sheet = $workbook->addWorksheet("Example");
        $sheet->write(0, 0, "Example");

        $this->assertSameAsInFixture('example.xls', $workbook);
    }

}
