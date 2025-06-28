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
        $this->assertEquals(0x6F, $ptg['ptgMemNoMemN'], 
            'ptgMemNoMemN should be 0x6F (the last duplicate), not 0x2F or 0x4F');
        
        // Test ptgAreaErr3d - should have the LAST duplicate value  
        // Original duplicates: 0x3D (commented), 0x5D (commented), 0x7D (active)
        $this->assertArrayHasKey('ptgAreaErr3d', $ptg,
            'ptgAreaErr3d key should exist in ptg array');
        $this->assertEquals(0x7D, $ptg['ptgAreaErr3d'], 
            'ptgAreaErr3d should be 0x7D (the last duplicate), not 0x3D or 0x5D');
        
        // Verify that the specific duplicated keys exist only once
        // Count occurrences of keys starting with the duplicated names
        $ptgMemNoMemCount = 0;
        $ptgMemNoMemNCount = 0;
        $ptgAreaErr3dCount = 0;
        foreach (array_keys($ptg) as $key) {
            if (strpos($key, 'ptgMemNoMem') === 0) {
                $ptgMemNoMemCount++;
            }
            if ($key === 'ptgMemNoMemN') {
                $ptgMemNoMemNCount++;
            }
            if ($key === 'ptgAreaErr3d') {
                $ptgAreaErr3dCount++;
            }
        }
        
        // Should have exactly 2 ptgMemNoMem* keys: ptgMemNoMem and ptgMemNoMemN
        $this->assertEquals(2, $ptgMemNoMemCount, 
            'There should be exactly 2 ptgMemNoMem* keys: ptgMemNoMem and ptgMemNoMemN');
        $this->assertEquals(1, $ptgMemNoMemNCount, 
            'There should be exactly one ptgMemNoMemN key');
        $this->assertEquals(1, $ptgAreaErr3dCount,
            'There should be exactly one ptgAreaErr3d key');
        
        // Verify that ptgMemNoMem exists with value 0x28
        // (The duplicates at 0x48 and 0x68 are commented out per Excel spec)
        $this->assertEquals(0x28, $ptg['ptgMemNoMem'], 'ptgMemNoMem should be 0x28');
        
        // Verify the incorrectly named variants don't exist
        $this->assertArrayNotHasKey('ptgMemNoMemV', $ptg, 
            'ptgMemNoMemV should not exist (Excel spec calls it ptgMemNoMem)');
        $this->assertArrayNotHasKey('ptgMemNoMemA', $ptg,
            'ptgMemNoMemA should not exist (Excel spec calls it ptgMemNoMem)');
    }
}