<?php

namespace Test\Spreadsheet\Excel\Writer;

use Spreadsheet_Excel_Writer_Parser;

use function is_string;

class ParserTest extends \LegacyPHPUnit\TestCase
{
    /**
     * Test that _convertFunction returns a value for all code paths
     */
    public function testConvertFunctionReturnsValue()
    {
        $parser = new Spreadsheet_Excel_Writer_Parser(0, 0x0500);

        // Access protected method via reflection
        $method = new \ReflectionMethod($parser, '_convertFunction');
        $method->setAccessible(true);

        // Test with a function that has fixed args (should return early)
        // TIME has 3 fixed arguments
        $result = $method->invoke($parser, 'TIME', 3);
        $this->assertNotEmpty($result);
        $this->assertTrue(is_string($result));

        // Test variable args path - SUM has variable args
        $result = $method->invoke($parser, 'SUM', 2);
        $this->assertNotEmpty($result);
        $this->assertTrue(is_string($result));

        // Test that invalid argument counts throw an exception
        // Create a function with invalid args value
        // Array structure: [function_number, arg_count, unknown, volatile_flag]
        $parser->_functions['INVALID'] = array(999, -2, 0, 0); // -2 is not valid

        $this->expectException(\InvalidArgumentException::class);
        $this->expectExceptionMessage('Invalid argument count -2 for function INVALID');
        $method->invoke($parser, 'INVALID', 1);
    }

    /**
     * Test that duplicate PTG entries have the correct final values
     * This ensures backward compatibility is maintained
     *
     * Background: In the original code, these PTG names were duplicated with different values:
     * - ptgMemNoMemN appeared at 0x2F, 0x4F, and 0x6F
     * - ptgAreaErr3d appeared at 0x3D, 0x5D, and 0x7D
     *
     * In PHP arrays, duplicate keys result in the last value overwriting earlier ones.
     * This test confirms that behavior is preserved.
     */
    public function testDuplicatePtgValues()
    {
        $parser = new Spreadsheet_Excel_Writer_Parser(0, 0x0500);

        // Access protected property via reflection
        $property = new \ReflectionProperty($parser, 'ptg');
        $property->setAccessible(true);
        $ptg = $property->getValue($parser);

        // Test ptgMemNoMemN - should have the LAST duplicate value
        // Original duplicates: 0x2F (commented), 0x4F (commented), 0x6F (active)
        $this->assertArrayHasKey('ptgMemNoMemN', $ptg,
            'ptgMemNoMemN key should exist in ptg array');
        $this->assertSame(0x6F, $ptg['ptgMemNoMemN'],
            'ptgMemNoMemN should be 0x6F (the last duplicate), not 0x2F or 0x4F');

        // Test ptgAreaErr3d - should have the LAST duplicate value
        // Original duplicates: 0x3D (commented), 0x5D (commented), 0x7D (active)
        $this->assertArrayHasKey('ptgAreaErr3d', $ptg,
            'ptgAreaErr3d key should exist in ptg array');
        $this->assertSame(0x7D, $ptg['ptgAreaErr3d'],
            'ptgAreaErr3d should be 0x7D (the last duplicate), not 0x3D or 0x5D');

        // The assertArrayHasKey calls above already verify these keys exist exactly once
        // (PHP arrays cannot have duplicate keys)

        // Verify that ptgMemNoMem exists with value 0x28
        // (The duplicates at 0x48 and 0x68 are commented out per Excel spec)
        $this->assertSame(0x28, $ptg['ptgMemNoMem'], 'ptgMemNoMem should be 0x28');
    }
}
