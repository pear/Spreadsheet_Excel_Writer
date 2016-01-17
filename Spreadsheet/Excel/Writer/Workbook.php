<?php
/*
*  Module written/ported by Xavier Noguer <xnoguer@rezebra.com>
*
*  The majority of this is _NOT_ my code.  I simply ported it from the
*  PERL Spreadsheet::WriteExcel module.
*
*  The author of the Spreadsheet::WriteExcel module is John McNamara
*  <jmcnamara@cpan.org>
*
*  I _DO_ maintain this code, and John McNamara has nothing to do with the
*  porting of this code to PHP.  Any questions directly related to this
*  class library should be directed to me.
*
*  License Information:
*
*    Spreadsheet_Excel_Writer:  A library for generating Excel Spreadsheets
*    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*    This library is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU Lesser General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

require_once __DIR__ . '/Format.php';
require_once __DIR__ . '/BIFFwriter.php';
require_once __DIR__ . '/Worksheet.php';
require_once __DIR__ . '/Parser.php';

/**
* Class for generating Excel Spreadsheets
*
* @author   Xavier Noguer <xnoguer@rezebra.com>
* @category FileFormats
* @package  Spreadsheet_Excel_Writer
*/

class Spreadsheet_Excel_Writer_Workbook extends Spreadsheet_Excel_Writer_BIFFwriter
{
    /**
    * Filename for the Workbook
    * @var string
    */
    protected $fileName;

    /**
    * Formula parser
    * @var Spreadsheet_Excel_Writer_Parser
    */
    protected $parser;

    /**
    * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
    * @var integer
    */
    protected $flagFor1904;

    /**
    * The active worksheet of the workbook (0 indexed)
    * @var integer
    */
    protected $activeSheet;

    /**
    * 1st displayed worksheet in the workbook (0 indexed)
    * @var integer
    */
    protected $firstSheet;

    /**
    * Number of workbook tabs selected
    * @var integer
    */
    protected $selectedWorkBook;

    /**
    * Index for creating adding new formats to the workbook
    * @var integer
    */
    protected $xf_index;

    /**
    * Flag for preventing close from being called twice.
    * @var integer
    * @see close()
    */
    protected $fileIsClosed;

    /**
    * The BIFF file size for the workbook.
    * @var integer
    * @see calcSheetOffsets()
    */
    protected $biffSize;

    /**
    * The default sheetname for all sheets created.
    * @var string
    */
    protected $sheetName;

    /**
    * The default XF format.
    * @var Spreadsheet_Excel_Writer_Format
    */
    protected $temporaryFormat;

    /**
    * Array containing references to all of this workbook's worksheets
    * @var array
    */
    protected $workSheet;

    /**
    * Array of sheet names for creating the EXTERNSHEET records
    * @var array
    */
    protected $sheetNames;

    /**
    * Array containing references to all of this workbook's formats
    * @var array|Spreadsheet_Excel_Writer_Format[]
    */
    protected $formats;

    /**
    * Array containing the colour palette
    * @var array
    */
    protected $palette;

    /**
    * The default format for URLs.
    * @var Spreadsheet_Excel_Writer_Format
    */
    protected $urlFormat;

    /**
    * The codepage indicates the text encoding used for strings
    * @var integer
    */
    protected $codePage;

    /**
    * The country code used for localization
    * @var integer
    */
    protected $countryCode;

    /**
    * number of bytes for size info of strings
    * @var integer
    */
    protected $stringSizeInfoSize;

    /** @var int */
    protected $totalStringLength;

    /** @var int */
    protected $uniqueString;

    /** @var array */
    protected $tableOfStrings;

    /** @var int */
    protected $stringSizeInfo;

    /** @var array */
    private $blockSize;

    /**
    * Class constructor
    *
    * @param string $fileName for storing the workbook. "-" for writing to stdout.
    * @access public
    */
    public function __construct($fileName)
    {
        // It needs to call its parent's constructor explicitly
        parent::__construct();

        $this->fileName         = $fileName;
        $this->parser           = new Spreadsheet_Excel_Writer_Parser($this->byteOrder, $this->BIFF_version);
        $this->flagFor1904      = 0;
        $this->activeSheet      = 0;
        $this->firstSheet       = 0;
        $this->selectedWorkBook = 0;
        $this->xf_index         = 16; // 15 style XF's and 1 cell XF.
        $this->fileIsClosed       = 0;
        $this->biffSize         = 0;
        $this->sheetName        = 'Sheet';
        $this->temporaryFormat  = new Spreadsheet_Excel_Writer_Format($this->BIFF_version);
        $this->workSheet        = array();
        $this->sheetNames       = array();
        $this->formats          = array();
        $this->palette          = array();
        $this->codePage         = 0x04E4; // FIXME: should change for BIFF8
        $this->countryCode      = -1;
        $this->stringSizeInfo   = 3;

        // Add the default format for hyperlinks
        $this->urlFormat            = $this->addFormat(array('color' => 'blue', 'underline' => 1));
        $this->totalStringLength    = 0;
        $this->uniqueString         = 0;
        $this->tableOfStrings       = array();
        $this->setPaletteXl97();
    }

    /**
    * Calls finalization methods.
    * This method should always be the last one to be called on every workbook
    *
    * @access public
    * @return mixed true on success. PEAR_Error on failure
    */
    public function close()
    {
        if ($this->fileIsClosed) { // Prevent close() from being called twice.
            return true;
        }

        $resource = $this->storeWorkbook();

        if ($this->isError($resource)) {
            return $this->raiseError($resource->getMessage());
        }

        $this->fileIsClosed = 1;

        return true;
    }

    /**
    * An accessor for the workSheet[] array
    * Returns an array of the worksheet objects in a workbook
    * It actually calls to worksheets()
    *
    * @access public
    * @see worksheets()
    * @return array
    */
    public function sheets()
    {
        return $this->worksheets();
    }

    /**
    * An accessor for the workSheet[] array.
    * Returns an array of the worksheet objects in a workbook
    *
    * @access public
    * @return array
    */
    public function worksheets()
    {
        return $this->workSheet;
    }

    /**
    * Sets the BIFF version.
    * This method exists just to access experimental functionality
    * from BIFF8. It will be deprecated !
    * Only possible value is 8 (Excel 97/2000).
    * For any other value it fails silently.
    *
    * @access public
    * @param integer $version The BIFF version
    */
    public function setVersion($version)
    {
        if ($version == 8) { // only accept version 8
            $version = 0x0600;
            $this->BIFF_version = $version;
            // change BIFFwriter limit for CONTINUE records
            $this->limit = 8228;
            $this->temporaryFormat->_BIFF_version = $version;
            $this->urlFormat->_BIFF_version = $version;
            $this->parser->_BIFF_version = $version;
            $this->codePage = 0x04B0;

            $total_worksheets = count($this->workSheet);
            // change version for all worksheets too
            for ($i = 0; $i < $total_worksheets; ++$i) {
                $this->workSheet[$i]->_BIFF_version = $version;
            }

            $total_formats = count($this->formats);
            // change version for all formats too
            for ($i = 0; $i < $total_formats; ++$i) {
                $this->formats[$i]->_BIFF_version = $version;
            }
        }
    }

    /**
    * Set the country identifier for the workbook
    *
    * @access public
    * @param integer $code Is the international calling country code for the
    *                      chosen country.
    */
    public function setCountry($code)
    {
        $this->countryCode = $code;
    }

    /**
    * Add a new worksheet to the Excel workbook.
    * If no name is given the name of the worksheet will be Sheeti$i, with
    * $i in [1..].
    *
    * @access public
    * @param string $name the optional name of the worksheet
    * @return mixed reference to a worksheet object on success, PEAR_Error
    *               on failure
    */
    public function addWorksheet($name = '')
    {
        $index     = count($this->workSheet);
        $sheetName = $this->sheetName;

        if ($name == '') {
            $name = $sheetName . ($index+1);
        }

        // Check that sheetname is <= 31 chars (Excel limit before BIFF8).
        if ($this->BIFF_version != 0x0600) {
            if (strlen($name) > 31) {
                return $this->raiseError('Sheetname $name must be <= 31 chars');
            }
        } else {
            if (function_exists('iconv')) {
                $name = iconv('UTF-8','UTF-16LE',$name);
            }
        }

        // Check that the worksheet name doesn't already exist: a fatal Excel error.
        $total_worksheets = count($this->workSheet);

        for ($i = 0; $i < $total_worksheets; ++$i) {
            if ($this->workSheet[$i]->getName() == $name) {
                return $this->raiseError("Worksheet '$name' already exists");
            }
        }

        $worksheet = new Spreadsheet_Excel_Writer_Worksheet(
            $this->BIFF_version,
            $name,
            $index,
            $this->activeSheet,
            $this->firstSheet,
            $this->totalStringLength,
            $this->uniqueString,
            $this->tableOfStrings,
            $this->urlFormat,
            $this->parser,
            $this->temporaryDirectory
        );

        $this->workSheet[$index]    = $worksheet;    // Store ref for iterator
        $this->sheetNames[$index]   = $name;          // Store EXTERNSHEET names
        $this->parser->setExtSheet($name, $index);  // Register worksheet name with parser

        return $worksheet;
    }

    /**
     * Add a new format to the Excel workbook.
     * Also, pass any properties to the Format constructor.
     *
     * @access public
     * @param array $properties array with properties for initializing the format.
     * @return Spreadsheet_Excel_Writer_Format Spreadsheet_Excel_Writer_Format reference to an Excel Format
     */
    public function addFormat($properties = array())
    {
        $format = new Spreadsheet_Excel_Writer_Format(
            $this->BIFF_version,
            $this->xf_index,
            $properties
        );

        $this->xf_index    += 1;
        $this->formats[]    = $format;

        return $format;
    }

    /**
     * Create new validator.
     *
     * @access public
     * @return Spreadsheet_Excel_Writer_Validator reference to a Validator
     */
    public function addValidator()
    {
        /* FIXME: check for successful inclusion*/
        $valid = new Spreadsheet_Excel_Writer_Validator($this->parser);

        return $valid;
    }

    /**
    * Change the RGB components of the elements in the colour palette.
    *
    * @access public
    * @param integer $index colour index
    * @param integer $red   red RGB value [0-255]
    * @param integer $green green RGB value [0-255]
    * @param integer $blue  blue RGB value [0-255]
    * @return integer The palette index for the custom color
    */
    public function setCustomColor($index, $red, $green, $blue)
    {
        // Match a HTML #xxyyzz style parameter
        /*if (defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
            @_ = ($_[0], hex $1, hex $2, hex $3);
        }*/

        // Check that the colour index is the right range
        if ($index < 8 or $index > 64) {
            // TODO: assign real error codes
            return $this->raiseError('Color index $index outside range: 8 <= index <= 64');
        }

        // Check that the colour components are in the right range
        if (($red   < 0 or $red   > 255) ||
            ($green < 0 or $green > 255) ||
            ($blue  < 0 or $blue  > 255))
        {
            return $this->raiseError('Color component outside range: 0 <= color <= 255');
        }

        $index -= 8; // Adjust colour index (wingless dragonfly)

        // Set the RGB value
        $this->palette[$index] = array($red, $green, $blue, 0);

        return ($index + 8);
    }

    /**
    * Sets the colour palette to the Excel 97+ default.
    *
    * @access private
    */
    protected function setPaletteXl97()
    {
        $this->palette = array(
            array(0x00, 0x00, 0x00, 0x00),   // 8
            array(0xff, 0xff, 0xff, 0x00),   // 9
            array(0xff, 0x00, 0x00, 0x00),   // 10
            array(0x00, 0xff, 0x00, 0x00),   // 11
            array(0x00, 0x00, 0xff, 0x00),   // 12
            array(0xff, 0xff, 0x00, 0x00),   // 13
            array(0xff, 0x00, 0xff, 0x00),   // 14
            array(0x00, 0xff, 0xff, 0x00),   // 15
            array(0x80, 0x00, 0x00, 0x00),   // 16
            array(0x00, 0x80, 0x00, 0x00),   // 17
            array(0x00, 0x00, 0x80, 0x00),   // 18
            array(0x80, 0x80, 0x00, 0x00),   // 19
            array(0x80, 0x00, 0x80, 0x00),   // 20
            array(0x00, 0x80, 0x80, 0x00),   // 21
            array(0xc0, 0xc0, 0xc0, 0x00),   // 22
            array(0x80, 0x80, 0x80, 0x00),   // 23
            array(0x99, 0x99, 0xff, 0x00),   // 24
            array(0x99, 0x33, 0x66, 0x00),   // 25
            array(0xff, 0xff, 0xcc, 0x00),   // 26
            array(0xcc, 0xff, 0xff, 0x00),   // 27
            array(0x66, 0x00, 0x66, 0x00),   // 28
            array(0xff, 0x80, 0x80, 0x00),   // 29
            array(0x00, 0x66, 0xcc, 0x00),   // 30
            array(0xcc, 0xcc, 0xff, 0x00),   // 31
            array(0x00, 0x00, 0x80, 0x00),   // 32
            array(0xff, 0x00, 0xff, 0x00),   // 33
            array(0xff, 0xff, 0x00, 0x00),   // 34
            array(0x00, 0xff, 0xff, 0x00),   // 35
            array(0x80, 0x00, 0x80, 0x00),   // 36
            array(0x80, 0x00, 0x00, 0x00),   // 37
            array(0x00, 0x80, 0x80, 0x00),   // 38
            array(0x00, 0x00, 0xff, 0x00),   // 39
            array(0x00, 0xcc, 0xff, 0x00),   // 40
            array(0xcc, 0xff, 0xff, 0x00),   // 41
            array(0xcc, 0xff, 0xcc, 0x00),   // 42
            array(0xff, 0xff, 0x99, 0x00),   // 43
            array(0x99, 0xcc, 0xff, 0x00),   // 44
            array(0xff, 0x99, 0xcc, 0x00),   // 45
            array(0xcc, 0x99, 0xff, 0x00),   // 46
            array(0xff, 0xcc, 0x99, 0x00),   // 47
            array(0x33, 0x66, 0xff, 0x00),   // 48
            array(0x33, 0xcc, 0xcc, 0x00),   // 49
            array(0x99, 0xcc, 0x00, 0x00),   // 50
            array(0xff, 0xcc, 0x00, 0x00),   // 51
            array(0xff, 0x99, 0x00, 0x00),   // 52
            array(0xff, 0x66, 0x00, 0x00),   // 53
            array(0x66, 0x66, 0x99, 0x00),   // 54
            array(0x96, 0x96, 0x96, 0x00),   // 55
            array(0x00, 0x33, 0x66, 0x00),   // 56
            array(0x33, 0x99, 0x66, 0x00),   // 57
            array(0x00, 0x33, 0x00, 0x00),   // 58
            array(0x33, 0x33, 0x00, 0x00),   // 59
            array(0x99, 0x33, 0x00, 0x00),   // 60
            array(0x99, 0x33, 0x66, 0x00),   // 61
            array(0x33, 0x33, 0x99, 0x00),   // 62
            array(0x33, 0x33, 0x33, 0x00),   // 63
        );
    }

    /**
    * Assemble worksheets into a workbook and send the BIFF data to an OLE
    * storage.
    *
    * @access private
    * @return mixed true on success. PEAR_Error on failure
    */
    protected function storeWorkbook()
    {
        if (count($this->workSheet) == 0) {
            return true;
        }

        // Ensure that at least one worksheet has been selected.
        if ($this->activeSheet == 0) {
            $this->workSheet[0]->selected = 1;
        }

        // Calculate the number of selected worksheet tabs and call the finalization
        // methods for each worksheet
        $totalNumberOfWorkSheets = count($this->workSheet);

        for ($i = 0; $i < $totalNumberOfWorkSheets; ++$i) {
            if ($this->workSheet[$i]->selected) {
                ++$this->selectedWorkBook;
            }

            $this->workSheet[$i]->close($this->sheetNames);
        }

        // Add Workbook globals
        $this->storeBof(0x0005);
        $this->storeCodepage();

        if ($this->BIFF_version == 0x0600) {
            $this->storeWindow1();
        }

        if ($this->BIFF_version == 0x0500) {
            $this->storeExterns();    // For print area and repeat rows
        }

        $this->storeNames();      // For print area and repeat rows

        if ($this->BIFF_version == 0x0500) {
            $this->storeWindow1();
        }

        $this->storeDatemode();
        $this->storeAllFonts();
        $this->storeAllNumFormats();
        $this->storeAllXfs();
        $this->storeAllStyles();
        $this->storePalette();
        $this->calcSheetOffsets();

        // Add BOUNDSHEET records
        for ($i = 0; $i < $totalNumberOfWorkSheets; ++$i) {
            $this->storeBoundsheet($this->workSheet[$i]->name,$this->workSheet[$i]->offset);
        }

        if ($this->countryCode != -1) {
            $this->storeCountry();
        }

        if ($this->BIFF_version == 0x0600) {
            //$this->storeSupbookInternal();
            /* TODO: store external SUPBOOK records and XCT and CRN records
            in case of external references for BIFF8 */
            //$this->storeExternsheetBiff8();
            $this->storeSharedStringsTable();
        }

        // End Workbook globals
        $this->storeEof();

        // Store the workbook in an OLE container
        $res = $this->storeOLEFile();

        if ($this->isError($res)) {
            return $this->raiseError($res->getMessage());
        }

        return true;
    }

    /**
    * Store the workbook in an OLE container
    *
    * @access private
    * @return mixed true on success. PEAR_Error on failure
    */
    protected function storeOLEFile()
    {
        if ($this->BIFF_version == 0x0600) {
            $OLE = new OLE_PPS_File(OLE::Asc2Ucs('Workbook'));
        } else {
            $OLE = new OLE_PPS_File(OLE::Asc2Ucs('Book'));
        }

        if ($this->temporaryDirectory != '') {
            $OLE->setTempDir($this->temporaryDirectory);
        }

        $res = $OLE->init();

        if ($this->isError($res)) {
            return $this->raiseError('OLE Error: ' . $res->getMessage());
        }

        $OLE->append($this->data);

        $total_worksheets = count($this->workSheet);

        for ($i = 0; $i < $total_worksheets; ++$i) {
            while ($tmp = $this->workSheet[$i]->getData()) {
                $OLE->append($tmp);
            }
        }

        $root = new OLE_PPS_Root(time(), time(), array($OLE));

        if ($this->temporaryDirectory != '') {
            $root->setTempDir($this->temporaryDirectory);
        }

        $res = $root->save($this->fileName);

        if ($this->isError($res)) {
            return $this->raiseError('OLE Error: ' . $res->getMessage());
        }

        return true;
    }

    /**
    * Calculate offsets for Worksheet BOF records.
    *
    * @access private
    */
    protected function calcSheetOffsets()
    {
        if ($this->BIFF_version == 0x0600) {
            $lengthOfTheBoundSheet = 12;  // fixed length for a BOUNDSHEET record
        } else {
            $lengthOfTheBoundSheet = 11;
        }

        $EOF    = 4;
        $offset = $this->dataSize;

        if ($this->BIFF_version == 0x0600) {
            // add the length of the SST
            /* TODO: check this works for a lot of strings (> 8224 bytes) */
            $offset += $this->calculateSharedStringsSizes();

            if ($this->countryCode != -1) {
                $offset += 8; // adding COUNTRY record
            }
            // add the lenght of SUPBOOK, EXTERNSHEET and NAME records
            //$offset += 8; // FIXME: calculate real value when storing the records
        }
        $totalNumberOfWorkSheets = count($this->workSheet);
        // add the length of the BOUNDSHEET records

        for ($i = 0; $i < $totalNumberOfWorkSheets; ++$i) {
            $offset += $lengthOfTheBoundSheet + strlen($this->workSheet[$i]->name);
        }
        $offset += $EOF;

        for ($i = 0; $i < $totalNumberOfWorkSheets; ++$i) {
            $this->workSheet[$i]->offset = $offset;
            $offset += $this->workSheet[$i]->_datasize;
        }

        $this->biffSize = $offset;
    }

    /**
    * Store the Excel FONT records.
    *
    * @access private
    */
    protected function storeAllFonts()
    {
        // tmp_format is added by the constructor. We use this to write the default XF's
        $format = $this->temporaryFormat;
        $font   = $format->getFont();

        // Note: Fonts are 0-indexed. According to the SDK there is no index 4,
        // so the following fonts are 0, 1, 2, 3, 5
        //
        for ($i = 1; $i <= 5; $i++){
            $this->append($font);
        }

        // Iterate through the XF objects and write a FONT record if it isn't the
        // same as the default FONT and if it hasn't already been used.
        //
        $fonts = array();
        $index = 6;                  // The first user defined FONT

        $key = $format->getFontKey(); // The default font from temporaryFormat
        $fonts[$key] = 0;             // Index of the default font

        $totalNumberOfFormats = count($this->formats);

        for ($i = 0; $i < $totalNumberOfFormats; ++$i) {
            $key = $this->formats[$i]->getFontKey();

            if (isset($fonts[$key])) {
                // FONT has already been used
                $this->formats[$i]->font_index = $fonts[$key];
            } else {
                // Add a new FONT record
                $fonts[$key]        = $index;
                $this->formats[$i]->font_index = $index;
                ++$index;
                $font = $this->formats[$i]->getFont();
                $this->append($font);
            }
        }
    }

    /**
    * Store user defined numerical formats i.e. FORMAT records
    *
    * @access private
    */
    protected function storeAllNumFormats()
    {
        // Leaning num_format syndrome
        $index              = 164;
        $numberFormatHashes = array();
        $numberFormats      = array();

        // Iterate through the XF objects and write a FORMAT record if it isn't a
        // built-in format type and if the FORMAT string hasn't already been used.
        $total_formats = count($this->formats);

        for ($i = 0; $i < $total_formats; ++$i) {
            $numberFormat = $this->formats[$i]->_num_format;

            // Check if $num_format is an index to a built-in format.
            // Also check for a string of zeros, which is a valid format string
            // but would evaluate to zero.
            //
            //@todo check if we can saftly replace >>"<< with >>'<<
            if (!preg_match("/^0+\d/", $numberFormat)) {
                if (preg_match("/^\d+$/", $numberFormat)) { // built-in format
                    continue;
                }
            }

            if (isset($numberFormatHashes[$numberFormat])) {
                // FORMAT has already been used
                $this->formats[$i]->_num_format = $numberFormatHashes[$numberFormat];
            } else{
                // Add a new FORMAT
                $numberFormatHashes[$numberFormat]    = $index;
                $this->formats[$i]->_num_format     = $index;
                array_push($numberFormats,$numberFormat);
                ++$index;
            }
        }

        // Write the new FORMAT records starting from 0xA4
        $index = 164;
        foreach ($numberFormats as $numberFormat) {
            $this->storeNumFormat($numberFormat,$index);
            ++$index;
        }
    }

    /**
    * Write all XF records.
    *
    * @access private
    */
    public function storeAllXfs()
    {
        // temporaryFormat is added by the constructor. We use this to write the default XF's
        // The default font index is 0
        //
        $format = $this->temporaryFormat;

        for ($i = 0; $i <= 14; ++$i) {
            $xf = $format->getXf('style'); // Style XF
            $this->append($xf);
        }

        $xf = $format->getXf('cell');      // Cell XF
        $this->append($xf);

        // User defined XFs
        $total_formats = count($this->formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $xf = $this->formats[$i]->getXf('cell');
            $this->append($xf);
        }
    }

    /**
    * Write all STYLE records.
    *
    * @access private
    */
    protected function storeAllStyles()
    {
        $this->storeStyle();
    }

    /**
    * Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
    * the NAME records.
    *
    * @access private
    */
    protected function storeExterns()
    {
        // Create EXTERNCOUNT with number of worksheets
        $this->storeExterncount(count($this->workSheet));

        // Create EXTERNSHEET for each worksheet
        foreach ($this->sheetNames as $sheetname) {
            $this->storeExternsheet($sheetname);
        }
    }

    /**
    * Write the NAME record to define the print area and the repeat rows and cols.
    *
    * @access private
    */
    protected function storeNames()
    {
        // Create the print area NAME records
        $total_worksheets = count($this->workSheet);

        for ($i = 0; $i < $total_worksheets; ++$i) {
            // Write a Name record if the print area has been defined
            if (isset($this->workSheet[$i]->print_rowmin)) {
                $this->storeNameShort(
                    $this->workSheet[$i]->index,
                    0x06, // NAME type
                    $this->workSheet[$i]->print_rowmin,
                    $this->workSheet[$i]->print_rowmax,
                    $this->workSheet[$i]->print_colmin,
                    $this->workSheet[$i]->print_colmax
                    );
            }
        }

        // Create the print title NAME records
        $total_worksheets = count($this->workSheet);

        for ($i = 0; $i < $total_worksheets; ++$i) {
            $rowmin = $this->workSheet[$i]->title_rowmin;
            $rowmax = $this->workSheet[$i]->title_rowmax;
            $colmin = $this->workSheet[$i]->title_colmin;
            $colmax = $this->workSheet[$i]->title_colmax;

            // Determine if row + col, row, col or nothing has been defined
            // and write the appropriate record
            //
            if (isset($rowmin) && isset($colmin)) {
                // Row and column titles have been defined.
                // Row title has been defined.
                $this->storeNameLong(
                    $this->workSheet[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    $colmin,
                    $colmax
                    );
            } elseif (isset($rowmin)) {
                // Row title has been defined.
                $this->storeNameShort(
                    $this->workSheet[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    0x00,
                    0xff
                    );
            } elseif (isset($colmin)) {
                // Column title has been defined.
                $this->storeNameShort(
                    $this->workSheet[$i]->index,
                    0x07, // NAME type
                    0x0000,
                    0x3fff,
                    $colmin,
                    $colmax
                    );
            } else {
                // Print title hasn't been defined.
            }
        }
    }




    /******************************************************************************
    *
    * BIFF RECORDS
    *
    */

    /**
    * Stores the CODEPAGE biff record.
    *
    * @access private
    */
    protected function storeCodepage()
    {
        $record          = 0x0042;             // Record identifier
        $length          = 0x0002;             // Number of bytes to follow
        $cv              = $this->codePage;   // The code page

        $header          = pack('vv', $record, $length);
        $data            = pack('v',  $cv);

        $this->append($header . $data);
    }

    /**
    * Write Excel BIFF WINDOW1 record.
    *
    * @access private
    */
    protected function storeWindow1()
    {
        $record    = 0x003D;                 // Record identifier
        $length    = 0x0012;                 // Number of bytes to follow

        $xWn       = 0x0000;                 // Horizontal position of window
        $yWn       = 0x0000;                 // Vertical position of window
        $dxWn      = 0x25BC;                 // Width of window
        $dyWn      = 0x1572;                 // Height of window

        $grbit     = 0x0038;                 // Option flags
        $ctabsel   = $this->selectedWorkBook;       // Number of workbook tabs selected
        $wTabRatio = 0x0258;                 // Tab to scrollbar ratio

        $itabFirst = $this->firstSheet;     // 1st displayed worksheet
        $itabCur   = $this->activeSheet;    // Active worksheet

        $header    = pack('vv', $record, $length);
        $data      = pack('vvvvvvvvv', $xWn, $yWn, $dxWn, $dyWn,
                                       $grbit,
                                       $itabCur, $itabFirst,
                                       $ctabsel, $wTabRatio);
        $this->append($header . $data);
    }

    /**
    * Writes Excel BIFF BOUNDSHEET record.
    * FIXME: inconsistent with BIFF documentation
    *
    * @param string  $sheetName Worksheet name
    * @param integer $offset    Location of worksheet BOF
    * @access private
    */
    protected function storeBoundsheet($sheetName, $offset)
    {
        $record = 0x0085;                    // Record identifier

        if ($this->BIFF_version == 0x0600) {
            $length = 0x08 + strlen($sheetName); // Number of bytes to follow
        } else {
            $length = 0x07 + strlen($sheetName); // Number of bytes to follow
        }

        $grbit = 0x0000;                    // Visibility and sheet type

        if ($this->BIFF_version == 0x0600) {
            $cch       = mb_strlen($sheetName,'UTF-16LE'); // Length of sheet name
        } else {
            $cch       = strlen($sheetName);        // Length of sheet name
        }

        $header = pack('vv',  $record, $length);

        if ($this->BIFF_version == 0x0600) {
            $data      = pack('VvCC', $offset, $grbit, $cch, 0x1);
        } else {
            $data      = pack('VvC', $offset, $grbit, $cch);
        }

        $this->append($header . $data . $sheetName);
    }

    /**
    * Write Internal SUPBOOK record
    *
    * @access private
    */
    protected function storeSupbookInternal()
    {
        $record    = 0x01AE;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $header    = pack('vv', $record, $length);
        $data      = pack('vv', count($this->workSheet), 0x0104);
        $this->append($header . $data);
    }

    /**
    * Writes the Excel BIFF EXTERNSHEET record. These references are used by
    * formulas.
    *
    * @param string $sheetname Worksheet name
    * @access private
    */
    protected function storeExternsheetBiff8()
    {
        $total_references = count($this->parser->_references);
        $record = 0x0017;                     // Record identifier
        $length = 2 + 6 * $total_references;  // Number of bytes to follow
        $header = pack('vv',  $record, $length);
        $data   = pack('v', $total_references);

        for ($i = 0; $i < $total_references; ++$i) {
            $data .= $this->parser->_references[$i];
        }

        $this->append($header . $data);
    }

    /**
    * Write Excel BIFF STYLE records.
    *
    * @access private
    */
    protected function storeStyle()
    {
        $record    = 0x0293;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $ixfe      = 0x8000;   // Index to style XF
        $BuiltIn   = 0x00;     // Built-in style
        $iLevel    = 0xff;     // Outline style level

        $header    = pack('vv',  $record, $length);
        $data      = pack('vCC', $ixfe, $BuiltIn, $iLevel);
        $this->append($header . $data);
    }


    /**
    * Writes Excel FORMAT record for non "built-in" numerical formats.
    *
    * @param string  $format Custom format string
    * @param integer $ifmt   Format index code
    * @access private
    */
    protected function storeNumFormat($format, $ifmt)
    {
        $record    = 0x041E;                      // Record identifier

        if ($this->BIFF_version == 0x0600) {
            $length    = 5 + strlen($format);      // Number of bytes to follow
            $encoding = 0x0;
        } elseif ($this->BIFF_version == 0x0500) {
            $length    = 3 + strlen($format);      // Number of bytes to follow
        }

        if ( $this->BIFF_version == 0x0600 && function_exists('iconv') ) {     // Encode format String
            if (mb_detect_encoding($format, 'auto') !== 'UTF-16LE'){
                $format = iconv(mb_detect_encoding($format, 'auto'),'UTF-16LE',$format);
            }
            $encoding = 1;
            $cch = function_exists('mb_strlen') ? mb_strlen($format, 'UTF-16LE') : (strlen($format) / 2);
        } else {
            $encoding = 0;
            $cch  = strlen($format);             // Length of format string
        }
        $length = strlen($format);

        if ($this->BIFF_version == 0x0600) {
            $header    = pack('vv', $record, 5 + $length);
            $data      = pack('vvC', $ifmt, $cch, $encoding);
        } elseif ($this->BIFF_version == 0x0500) {
            $header    = pack('vv', $record, 3 + $length);
            $data      = pack('vC', $ifmt, $cch);
        }
        $this->append($header . $data . $format);
    }

    /**
    * Write DATEMODE record to indicate the date system in use (1904 or 1900).
    *
    * @access private
    */
    protected function storeDatemode()
    {
        $record    = 0x0022;         // Record identifier
        $length    = 0x0002;         // Bytes to follow

        $f1904     = $this->flagFor1904;   // Flag for 1904 date system

        $header    = pack('vv', $record, $length);
        $data      = pack('v', $f1904);
        $this->append($header . $data);
    }


    /**
    * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
    * references in the workbook.
    *
    * Excel only stores references to external sheets that are used in NAME.
    * The workbook NAME record is required to define the print area and the repeat
    * rows and columns.
    *
    * A similar method is used in Worksheet.php for a slightly different purpose.
    *
    * @param integer $cxals Number of external references
    * @access private
    */
    protected function storeExterncount($cxals)
    {
        $record   = 0x0016;          // Record identifier
        $length   = 0x0002;          // Number of bytes to follow

        $header   = pack('vv', $record, $length);
        $data     = pack('v',  $cxals);
        $this->append($header . $data);
    }


    /**
    * Writes the Excel BIFF EXTERNSHEET record. These references are used by
    * formulas. NAME record is required to define the print area and the repeat
    * rows and columns.
    *
    * A similar method is used in Worksheet.php for a slightly different purpose.
    *
    * @param string $sheetName Worksheet name
    * @access private
    */
    protected function storeExternsheet($sheetName)
    {
        $record      = 0x0017;                     // Record identifier
        $length      = 0x02 + strlen($sheetName);  // Number of bytes to follow

        $cch         = strlen($sheetName);         // Length of sheet name
        $rgch        = 0x03;                       // Filename encoding

        $header      = pack('vv',  $record, $length);
        $data        = pack('CC', $cch, $rgch);
        $this->append($header . $data . $sheetName);
    }


    /**
    * Store the NAME record in the short format that is used for storing the print
    * area, repeat rows only and repeat columns only.
    *
    * @param integer $index  Sheet index
    * @param integer $type   Built-in name type
    * @param integer $rowmin Start row
    * @param integer $rowmax End row
    * @param integer $colmin Start colum
    * @param integer $colmax End column
    * @access private
    */
    protected function storeNameShort($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x0024;       // Number of bytes to follow

        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x0015;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown03       = 0x3b;
        $unknown04       = 0xffff-$index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8005;

        $header             = pack('vv', $record, $length);
        $data               = pack('v', $grbit);
        $data              .= pack('C', $chKey);
        $data              .= pack('C', $cch);
        $data              .= pack('v', $cce);
        $data              .= pack('v', $ixals);
        $data              .= pack('v', $itab);
        $data              .= pack('C', $cchCustMenu);
        $data              .= pack('C', $cchDescription);
        $data              .= pack('C', $cchHelptopic);
        $data              .= pack('C', $cchStatustext);
        $data              .= pack('C', $rgch);
        $data              .= pack('C', $unknown03);
        $data              .= pack('v', $unknown04);
        $data              .= pack('v', $unknown05);
        $data              .= pack('v', $unknown06);
        $data              .= pack('v', $unknown07);
        $data              .= pack('v', $unknown08);
        $data              .= pack('v', $index);
        $data              .= pack('v', $index);
        $data              .= pack('v', $rowmin);
        $data              .= pack('v', $rowmax);
        $data              .= pack('C', $colmin);
        $data              .= pack('C', $colmax);
        $this->append($header . $data);
    }


    /**
    * Store the NAME record in the long format that is used for storing the repeat
    * rows and columns when both are specified. This shares a lot of code with
    * storeNameShort() but we use a separate method to keep the code clean.
    * Code abstraction for reuse can be carried too far, and I should know. ;-)
    *
    * @param integer $index Sheet index
    * @param integer $type  Built-in name type
    * @param integer $rowmin Start row
    * @param integer $rowmax End row
    * @param integer $colmin Start colum
    * @param integer $colmax End column
    * @access private
    */
    protected function storeNameLong($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x003d;       // Number of bytes to follow
        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x002e;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown01       = 0x29;
        $unknown02       = 0x002b;
        $unknown03       = 0x3b;
        $unknown04       = 0xffff-$index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8008;

        $header             = pack('vv',  $record, $length);
        $data               = pack('v', $grbit);
        $data              .= pack('C', $chKey);
        $data              .= pack('C', $cch);
        $data              .= pack('v', $cce);
        $data              .= pack('v', $ixals);
        $data              .= pack('v', $itab);
        $data              .= pack('C', $cchCustMenu);
        $data              .= pack('C', $cchDescription);
        $data              .= pack('C', $cchHelptopic);
        $data              .= pack('C', $cchStatustext);
        $data              .= pack('C', $rgch);
        $data              .= pack('C', $unknown01);
        $data              .= pack('v', $unknown02);
        // Column definition
        $data              .= pack('C', $unknown03);
        $data              .= pack('v', $unknown04);
        $data              .= pack('v', $unknown05);
        $data              .= pack('v', $unknown06);
        $data              .= pack('v', $unknown07);
        $data              .= pack('v', $unknown08);
        $data              .= pack('v', $index);
        $data              .= pack('v', $index);
        $data              .= pack('v', 0x0000);
        $data              .= pack('v', 0x3fff);
        $data              .= pack('C', $colmin);
        $data              .= pack('C', $colmax);
        // Row definition
        $data              .= pack('C', $unknown03);
        $data              .= pack('v', $unknown04);
        $data              .= pack('v', $unknown05);
        $data              .= pack('v', $unknown06);
        $data              .= pack('v', $unknown07);
        $data              .= pack('v', $unknown08);
        $data              .= pack('v', $index);
        $data              .= pack('v', $index);
        $data              .= pack('v', $rowmin);
        $data              .= pack('v', $rowmax);
        $data              .= pack('C', 0x00);
        $data              .= pack('C', 0xff);
        // End of data
        $data              .= pack('C', 0x10);
        $this->append($header . $data);
    }

    /**
    * Stores the COUNTRY record for localization
    *
    * @access private
    */
    protected function storeCountry()
    {
        $record          = 0x008C;    // Record identifier
        $length          = 4;         // Number of bytes to follow

        $header = pack('vv',  $record, $length);
        /* using the same country code always for simplicity */
        $data = pack('vv', $this->countryCode, $this->countryCode);
        $this->append($header . $data);
    }

    /**
    * Stores the PALETTE biff record.
    *
    * @access private
    */
    protected function storePalette()
    {
        $aref            = $this->palette;

        $record          = 0x0092;                 // Record identifier
        $length          = 2 + 4 * count($aref);   // Number of bytes to follow
        $ccv             =         count($aref);   // Number of RGB values to follow
        $data = '';                                // The RGB data

        // Pack the RGB data
        foreach ($aref as $color) {
            foreach ($color as $byte) {
                $data .= pack('C', $byte);
            }
        }

        $header = pack('vvv',  $record, $length, $ccv);
        $this->append($header . $data);
    }

    /**
    * Calculate
    * Handling of the SST continue blocks is complicated by the need to include an
    * additional continuation byte depending on whether the string is split between
    * blocks or whether it starts at the beginning of the block. (There are also
    * additional complications that will arise later when/if Rich Strings are
    * supported).
    *
    * @access private
    */
    protected function calculateSharedStringsSizes()
    {
        /* Iterate through the strings to calculate the CONTINUE block sizes.
           For simplicity we use the same size for the SST and CONTINUE records:
           8228 : Maximum Excel97 block size
             -4 : Length of block header
             -8 : Length of additional SST header information
             -8 : Arbitrary number to keep within _add_continue() limit = 8208
        */
        $continue_limit     = 8208;
        $block_length       = 0;
        $written            = 0;
        $this->blockSize = array();
        $continue           = 0;

        foreach (array_keys($this->tableOfStrings) as $string) {
            $string_length = strlen($string);
            $headerinfo    = unpack('vlength/Cencoding', $string);
            $encoding      = $headerinfo['encoding'];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $block_length += $string_length;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                $written      += $string_length;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            while ($block_length >= $continue_limit) {

                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;


                /* TODO: Unicode data should only be split on char (2 byte)
                boundaries. Therefore, in some cases we need to reduce the
                amount of available
                */
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$split_string && $space_remaining % 2 != 1) {
                            $space_remaining--;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        else if ($split_string && $space_remaining % 2 == 1) {
                            $space_remaining--;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }


                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $written      += $space_remaining;

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // Store the max size for this block
                    $this->blockSize[] = $continue_limit - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Store the max size for this block
                    $this->blockSize[] = $written + $continue;

                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;

                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                if ($block_length < $continue_limit) {
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }

        // Store the max size for the last block unless it is empty
        if ($written + $continue) {
            $this->blockSize[] = $written + $continue;
        }


        /* Calculate the total length of the SST and associated CONTINUEs (if any).
         The SST record will have a length even if it contains no strings.
         This length is required to set the offsets in the BOUNDSHEET records since
         they must be written before the SST records
        */

        $tmp_block_sizes    = $this->blockSize;
        $length             = 12;

        if (!empty($tmp_block_sizes)) {
            $length += array_shift($tmp_block_sizes); // SST
        }

        while (!empty($tmp_block_sizes)) {
            $length += 4 + array_shift($tmp_block_sizes); // CONTINUEs
        }

        return $length;
    }

    /**
    * Write all of the workbooks strings into an indexed array.
    * See the comments in _calculate_shared_string_sizes() for more information.
    *
    * The Excel documentation says that the SST record should be followed by an
    * EXTSST record. The EXTSST record is a hash table that is used to optimise
    * access to SST. However, despite the documentation it doesn't seem to be
    * required so we will ignore it.
    *
    * @access private
    */
    protected function storeSharedStringsTable()
    {
        $record  = 0x00fc;  // Record identifier
        $length  = 0x0008;  // Number of bytes to follow
        $total   = 0x0000;

        // Iterate through the strings to calculate the CONTINUE block sizes
        $continue_limit = 8208;
        $block_length   = 0;
        $written        = 0;
        $continue       = 0;

        // sizes are upside down
        $tmp_block_sizes = $this->blockSize;
        // $tmp_block_sizes = array_reverse($this->blockSize);

        // The SST record is required even if it contains no strings. Thus we will
        // always have a length
        //
        if (!empty($tmp_block_sizes)) {
            $length = 8 + array_shift($tmp_block_sizes);
        }
        else {
            // No strings
            $length = 8;
        }



        // Write the SST block header information
        $header      = pack('vv', $record, $length);
        $data        = pack('VV', $this->totalStringLength, $this->uniqueString);
        $this->append($header . $data);




        /* TODO: not good for performance */
        foreach (array_keys($this->tableOfStrings) as $string) {

            $string_length = strlen($string);
            $headerinfo    = unpack('vlength/Cencoding', $string);
            $encoding      = $headerinfo['encoding'];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            //
            $block_length += $string_length;


            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                $this->append($string);
                $written += $string_length;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            //
            while ($block_length >= $continue_limit) {

                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                //
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;


                // Unicode data should only be split on char (2 byte) boundaries.
                // Therefore, in some cases we need to reduce the amount of available
                // space by 1 byte to ensure the correct alignment.
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$split_string && $space_remaining % 2 != 1) {
                            $space_remaining--;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        else if ($split_string && $space_remaining % 2 == 1) {
                            $space_remaining--;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }


                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $tmp = substr($string, 0, $space_remaining);
                    $this->append($tmp);

                    // The remainder will be written in the next block(s)
                    $string = substr($string, $space_remaining);

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    //
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;
                }

                // Write the CONTINUE block header
                if (!empty($this->blockSize)) {
                    $record  = 0x003C;
                    $length  = array_shift($tmp_block_sizes);

                    $header  = pack('vv', $record, $length);
                    if ($continue) {
                        $header .= pack('C', $encoding);
                    }
                    $this->append($header);
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                //
                if ($block_length < $continue_limit) {
                    $this->append($string);
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }
    }
}

