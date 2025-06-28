<?php

namespace Test\Spreadsheet\Excel\Writer;

use Spreadsheet_Excel_Writer_Parser;

class ParserTest extends \LegacyPHPUnit\TestCase
{
    /**
     * Test that _convertFunction returns a value for all code paths
     */
    public function testConvertFunctionReturnsValue()
    {
        $parser = new Spreadsheet_Excel_Writer_Parser(0, 0x0500);

        // Initialize the parser properly
        $parser->_initializeHashes();

        // Access protected method via reflection
        $method = new \ReflectionMethod($parser, '_convertFunction');
        $method->setAccessible(true);

        // Test with a function that has fixed args (should return early)
        // TIME has 3 fixed arguments
        $result = $method->invoke($parser, 'TIME', 3);
        $this->assertNotEmpty($result);
        $this->assertInternalType('string', $result);

        // Test variable args path - SUM has variable args
        $result = $method->invoke($parser, 'SUM', 2);
        $this->assertNotEmpty($result);
        $this->assertInternalType('string', $result);
    }
}