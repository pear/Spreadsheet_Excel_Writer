<?php

namespace Test\Spreadsheet\Excel\Writer;

use Spreadsheet_Excel_Writer_Worksheet;
use Spreadsheet_Excel_Writer_Workbook;

class WorksheetTest extends \LegacyPHPUnit\TestCase
{
    private $workbook;
    private $worksheet;

    public function doSetUp()
    {
        parent::doSetUp();
        $this->workbook = new Spreadsheet_Excel_Writer_Workbook('php://memory');

        $activesheet = 0;
        $str_total = 0;
        $str_unique = 0;
        $str_table = 0;
        $url_format = '';
        $parser = '';
        $tmp_dir = '';

        $this->worksheet = new Spreadsheet_Excel_Writer_Worksheet(0x0500, 'Test', 0, $activesheet, $this->workbook->_url_format, $str_total, $str_unique, $str_table, $url_format, $parser, $tmp_dir);
    }

    public function doTearDown()
    {
        if ($this->workbook) {
            $this->workbook->close();
        }
        parent::doTearDown();
    }

    public function testIt()
    {
        $this->assertInstanceOf('Spreadsheet_Excel_Writer_Worksheet', $this->worksheet);
    }
}
