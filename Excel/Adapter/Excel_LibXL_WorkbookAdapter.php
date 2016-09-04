<?php

namespace MultiLibExcelExport\Excel\Adapter;

use MultiLibExcelExport\Excel\Adapter\Excel_Adapter;
use MultiLibExcelExport\Excel\Adapter\Excel_WorkbookAdapterInterface;

class Excel_LibXL_WorkbookAdapter extends Excel_Adapter implements Excel_WorkbookAdapterInterface
{
    static protected $_INCLUDES;
    
    static protected $_WORKSHEET_CLASS;
    static protected $_STYLE_CLASS;
    
    const LICENSE_NAME = '';
    const LICENSE_KEY = '';
    
    protected $_oExcelBook;
    protected $_sWorkbookFilename;
    
    protected $_fDefaultColSize;
    protected $_aColWidth;
    
    protected $_aPictureId;
    protected $_aCustomNumberFormat;
    
    //Constantes d'ExcelFont
    const ExcelFont_NORMAL = 0;
    const ExcelFont_SUPERSCRIPT = 1;
    const ExcelFont_SUBSCRIPT = 2;

    const ExcelFont_UNDERLINE_NONE = 0;
    const ExcelFont_UNDERLINE_SINGLE = 1;
    const ExcelFont_UNDERLINE_DOUBLE = 2;
    const ExcelFont_UNDERLINE_SINGLEACC = 33;
    const ExcelFont_UNDERLINE_DOUBLEACC = 34;
    
    //Constantes d'ExcelFormat
    const ExcelFormat_COLOR_BLACK = 8;
    const ExcelFormat_COLOR_WHITE = 9;
    const ExcelFormat_COLOR_RED = 10;
    const ExcelFormat_COLOR_BRIGHTGREEN = 11;
    const ExcelFormat_COLOR_BLUE = 12;
    const ExcelFormat_COLOR_YELLOW = 13;
    const ExcelFormat_COLOR_PINK = 14;
    const ExcelFormat_COLOR_TURQUOISE = 15;
    const ExcelFormat_COLOR_DARKRED = 16;
    const ExcelFormat_COLOR_GREEN = 17;
    const ExcelFormat_COLOR_DARKBLUE = 18;
    const ExcelFormat_COLOR_DARKYELLOW = 19;
    const ExcelFormat_COLOR_VIOLET = 20;
    const ExcelFormat_COLOR_TEAL = 21;
    const ExcelFormat_COLOR_GRAY25 = 22;
    const ExcelFormat_COLOR_GRAY50 = 23;
    const ExcelFormat_COLOR_PERIWINKLE_CF = 24;
    const ExcelFormat_COLOR_PLUM_CF = 25;
    const ExcelFormat_COLOR_IVORY_CF = 26;
    const ExcelFormat_COLOR_LIGHTTURQUOISE_CF = 27;
    const ExcelFormat_COLOR_DARKPURPLE_CF = 28;
    const ExcelFormat_COLOR_CORAL_CF = 29;
    const ExcelFormat_COLOR_OCEANBLUE_CF = 30;
    const ExcelFormat_COLOR_ICEBLUE_CF = 31;
    const ExcelFormat_COLOR_DARKBLUE_CL = 32;
    const ExcelFormat_COLOR_PINK_CL = 33;
    const ExcelFormat_COLOR_YELLOW_CL = 34;
    const ExcelFormat_COLOR_TURQUOISE_CL = 35;
    const ExcelFormat_COLOR_VIOLET_CL = 36;
    const ExcelFormat_COLOR_DARKRED_CL = 37;
    const ExcelFormat_COLOR_TEAL_CL = 38;
    const ExcelFormat_COLOR_BLUE_CL = 39;
    const ExcelFormat_COLOR_SKYBLUE = 40;
    const ExcelFormat_COLOR_LIGHTTURQUOISE = 41;
    const ExcelFormat_COLOR_LIGHTGREEN = 42;
    const ExcelFormat_COLOR_LIGHTYELLOW = 43;
    const ExcelFormat_COLOR_PALEBLUE = 44;
    const ExcelFormat_COLOR_ROSE = 45;
    const ExcelFormat_COLOR_LAVENDER = 46;
    const ExcelFormat_COLOR_TAN = 47;
    const ExcelFormat_COLOR_LIGHTBLUE = 48;
    const ExcelFormat_COLOR_AQUA = 49;
    const ExcelFormat_COLOR_LIME = 50;
    const ExcelFormat_COLOR_GOLD = 51;
    const ExcelFormat_COLOR_LIGHTORANGE = 52;
    const ExcelFormat_COLOR_ORANGE = 53;
    const ExcelFormat_COLOR_BLUEGRAY = 54;
    const ExcelFormat_COLOR_GRAY40 = 55;
    const ExcelFormat_COLOR_DARKTEAL = 56;
    const ExcelFormat_COLOR_SEAGREEN = 57;
    const ExcelFormat_COLOR_DARKGREEN = 58;
    const ExcelFormat_COLOR_OLIVEGREEN = 59;
    const ExcelFormat_COLOR_BROWN = 60;
    const ExcelFormat_COLOR_PLUM = 61;
    const ExcelFormat_COLOR_INDIGO = 62;
    const ExcelFormat_COLOR_GRAY80 = 63;
    const ExcelFormat_COLOR_DEFAULT_FOREGROUND = 64;
    const ExcelFormat_COLOR_DEFAULT_BACKGROUND = 65;
    const ExcelFormat_COLOR_TOOLTIP = 81;
    const ExcelFormat_COLOR_AUTO = 32767;
    
    const ExcelFormat_AS_DATE = 1;
    const ExcelFormat_AS_FORMULA = 2;
    const ExcelFormat_AS_NUMERIC_STRING = 3;
    
    const ExcelFormat_NUMFORMAT_GENERAL = 0;
    const ExcelFormat_NUMFORMAT_NUMBER = 1;
    const ExcelFormat_NUMFORMAT_NUMBER_D2 = 2;
    const ExcelFormat_NUMFORMAT_NUMBER_SEP = 3;
    const ExcelFormat_NUMFORMAT_NUMBER_SEP_D2 = 4;
    const ExcelFormat_NUMFORMAT_CURRENCY_NEGBRA = 5;
    const ExcelFormat_NUMFORMAT_CURRENCY_NEGBRARED = 6;
    const ExcelFormat_NUMFORMAT_CURRENCY_D2_NEGBRA = 7;
    const ExcelFormat_NUMFORMAT_CURRENCY_D2_NEGBRARED = 8;
    const ExcelFormat_NUMFORMAT_PERCENT = 9;
    const ExcelFormat_NUMFORMAT_PERCENT_D2 = 10;
    const ExcelFormat_NUMFORMAT_SCIENTIFIC_D2 = 11;
    const ExcelFormat_NUMFORMAT_FRACTION_ONEDIG = 12;
    const ExcelFormat_NUMFORMAT_FRACTION_TWODIG = 13;
    const ExcelFormat_NUMFORMAT_DATE = 14;
    const ExcelFormat_NUMFORMAT_CUSTOM_D_MON_YY = 15;
    const ExcelFormat_NUMFORMAT_CUSTOM_D_MON = 16;
    const ExcelFormat_NUMFORMAT_CUSTOM_MON_YY = 17;
    const ExcelFormat_NUMFORMAT_CUSTOM_HMM_AM = 18;
    const ExcelFormat_NUMFORMAT_CUSTOM_HMMSS_AM = 19;
    const ExcelFormat_NUMFORMAT_CUSTOM_HMM = 20;
    const ExcelFormat_NUMFORMAT_CUSTOM_HMMSS = 21;
    const ExcelFormat_NUMFORMAT_CUSTOM_MDYYYY_HMM = 22;
    const ExcelFormat_NUMFORMAT_NUMBER_SEP_NEGBRA = 37;
    const ExcelFormat_NUMFORMAT_NUMBER_SEP_NEGBRARED = 38;
    const ExcelFormat_NUMFORMAT_NUMBER_D2_SEP_NEGBRA = 39;
    const ExcelFormat_NUMFORMAT_NUMBER_D2_SEP_NEGBRARED = 40;
    const ExcelFormat_NUMFORMAT_ACCOUNT = 41;
    const ExcelFormat_NUMFORMAT_ACCOUNTCUR = 42;
    const ExcelFormat_NUMFORMAT_ACCOUNT_D2 = 43;
    const ExcelFormat_NUMFORMAT_ACCOUNT_D2_CUR = 44;
    const ExcelFormat_NUMFORMAT_CUSTOM_MMSS = 45;
    const ExcelFormat_NUMFORMAT_CUSTOM_H0MMSS = 46;
    const ExcelFormat_NUMFORMAT_CUSTOM_MMSS0 = 47;
    const ExcelFormat_NUMFORMAT_CUSTOM_000P0E_PLUS0 = 48;
    const ExcelFormat_NUMFORMAT_TEXT = 49;
    
    const ExcelFormat_ALIGNH_GENERAL = 0;
    const ExcelFormat_ALIGNH_LEFT = 1;
    const ExcelFormat_ALIGNH_CENTER = 2;
    const ExcelFormat_ALIGNH_RIGHT = 3;
    const ExcelFormat_ALIGNH_FILL = 4;
    const ExcelFormat_ALIGNH_JUSTIFY = 5;
    const ExcelFormat_ALIGNH_MERGE = 6;
    const ExcelFormat_ALIGNH_DISTRIBUTED = 7;
    
    const ExcelFormat_ALIGNV_TOP = 0;
    const ExcelFormat_ALIGNV_CENTER = 1;
    const ExcelFormat_ALIGNV_BOTTOM = 2;
    const ExcelFormat_ALIGNV_JUSTIFY = 3;
    const ExcelFormat_ALIGNV_DISTRIBUTED = 4;
    
    const ExcelFormat_BORDERSTYLE_NONE = 0;
    const ExcelFormat_BORDERSTYLE_THIN = 1;
    const ExcelFormat_BORDERSTYLE_MEDIUM = 2;
    const ExcelFormat_BORDERSTYLE_DASHED = 3;
    const ExcelFormat_BORDERSTYLE_DOTTED = 4;
    const ExcelFormat_BORDERSTYLE_THICK = 5;
    const ExcelFormat_BORDERSTYLE_DOUBLE = 6;
    const ExcelFormat_BORDERSTYLE_HAIR = 7;
    const ExcelFormat_BORDERSTYLE_MEDIUMDASHED = 8;
    const ExcelFormat_BORDERSTYLE_DASHDOT = 9;
    const ExcelFormat_BORDERSTYLE_MEDIUMDASHDOT = 10;
    const ExcelFormat_BORDERSTYLE_DASHDOTDOT = 11;
    const ExcelFormat_BORDERSTYLE_MEDIUMDASHDOTDOT = 12;
    const ExcelFormat_BORDERSTYLE_SLANTDASHDOT = 13;
    
    const ExcelFormat_BORDERDIAGONAL_NONE = 0;
    const ExcelFormat_BORDERDIAGONAL_DOWN = 1;
    const ExcelFormat_BORDERDIAGONAL_UP = 2;
    const ExcelFormat_BORDERDIAGONAL_BOTH = 3;
    
    const ExcelFormat_FILLPATTERN_NONE = 0;
    const ExcelFormat_FILLPATTERN_SOLID = 1;
    const ExcelFormat_FILLPATTERN_GRAY50 = 2;
    const ExcelFormat_FILLPATTERN_GRAY75 = 3;
    const ExcelFormat_FILLPATTERN_GRAY25 = 4;
    const ExcelFormat_FILLPATTERN_HORSTRIPE = 5;
    const ExcelFormat_FILLPATTERN_VERSTRIPE = 6;
    const ExcelFormat_FILLPATTERN_REVDIAGSTRIPE = 7;
    const ExcelFormat_FILLPATTERN_DIAGSTRIPE = 8;
    const ExcelFormat_FILLPATTERN_DIAGCROSSHATCH = 9;
    const ExcelFormat_FILLPATTERN_THICKDIAGCROSSHATCH = 10;
    const ExcelFormat_FILLPATTERN_THINHORSTRIPE = 11;
    const ExcelFormat_FILLPATTERN_THINVERSTRIPE = 12;
    const ExcelFormat_FILLPATTERN_THINREVDIAGSTRIPE = 13;
    const ExcelFormat_FILLPATTERN_THINDIAGSTRIPE = 14;
    const ExcelFormat_FILLPATTERN_THINHORCROSSHATCH = 15;
    const ExcelFormat_FILLPATTERN_THINDIAGCROSSHATCH = 16;
    const ExcelFormat_FILLPATTERN_GRAY12P5 = 17;
    const ExcelFormat_FILLPATTERN_GRAY6P25 = 18;
    
    //Constantes d'ExcelSheet
    const ExcelSheet_PAPER_DEFAULT = 0;
    const ExcelSheet_PAPER_LETTER = 1;
    const ExcelSheet_PAPER_LETTERSMALL = 2;
    const ExcelSheet_PAPER_TABLOID = 3;
    const ExcelSheet_PAPER_LEDGER = 4;
    const ExcelSheet_PAPER_LEGAL = 5;
    const ExcelSheet_PAPER_STATEMENT = 6;
    const ExcelSheet_PAPER_EXECUTIVE = 7;
    const ExcelSheet_PAPER_A3 = 8;
    const ExcelSheet_PAPER_A4 = 9;
    const ExcelSheet_PAPER_A4SMALL = 10;
    const ExcelSheet_PAPER_A5 = 11;
    const ExcelSheet_PAPER_B4 = 12;
    const ExcelSheet_PAPER_B5 = 13;
    const ExcelSheet_PAPER_FOLIO = 14;
    const ExcelSheet_PAPER_QUATRO = 15;
    const ExcelSheet_PAPER_10x14 = 16;
    const ExcelSheet_PAPER_10x17 = 17;
    const ExcelSheet_PAPER_NOTE = 18;
    const ExcelSheet_PAPER_ENVELOPE_9 = 19;
    const ExcelSheet_PAPER_ENVELOPE_10 = 20;
    const ExcelSheet_PAPER_ENVELOPE_11 = 21;
    const ExcelSheet_PAPER_ENVELOPE_12 = 22;
    const ExcelSheet_PAPER_ENVELOPE_14 = 23;
    const ExcelSheet_PAPER_C_SIZE = 24;
    const ExcelSheet_PAPER_D_SIZE = 25;
    const ExcelSheet_PAPER_E_SIZE = 26;
    const ExcelSheet_PAPER_ENVELOPE_DL = 27;
    const ExcelSheet_PAPER_ENVELOPE_C5 = 28;
    const ExcelSheet_PAPER_ENVELOPE_C3 = 29;
    const ExcelSheet_PAPER_ENVELOPE_C4 = 30;
    const ExcelSheet_PAPER_ENVELOPE_C6 = 31;
    const ExcelSheet_PAPER_ENVELOPE_C65 = 32;
    const ExcelSheet_PAPER_ENVELOPE_B4 = 33;
    const ExcelSheet_PAPER_ENVELOPE_B5 = 34;
    const ExcelSheet_PAPER_ENVELOPE_B6 = 35;
    const ExcelSheet_PAPER_ENVELOPE = 36;
    const ExcelSheet_PAPER_ENVELOPE_MONARCH = 37;
    const ExcelSheet_PAPER_US_ENVELOPE = 38;
    const ExcelSheet_PAPER_FANFOLD = 39;
    const ExcelSheet_PAPER_GERMAN_STD_FANFOLD = 40;
    const ExcelSheet_PAPER_GERMAN_LEGAL_FANFOLD = 41;
    const ExcelSheet_PAPER_B4_ISO = 42;
    const ExcelSheet_PAPER_JAPANESE_POSTCARD = 43;
    const ExcelSheet_PAPER_9x11 = 44;
    const ExcelSheet_PAPER_10x11 = 45;
    const ExcelSheet_PAPER_15x11 = 46;
    const ExcelSheet_PAPER_ENVELOPE_INVITE = 47;
    const ExcelSheet_PAPER_US_LETTER_EXTRA = 50;
    const ExcelSheet_PAPER_US_LEGAL_EXTRA = 51;
    const ExcelSheet_PAPER_US_TABLOID_EXTRA = 52;
    const ExcelSheet_PAPER_A4_EXTRA = 53;
    const ExcelSheet_PAPER_LETTER_TRANSVERSE = 54;
    const ExcelSheet_PAPER_A4_TRANSVERSE = 55;
    const ExcelSheet_PAPER_LETTER_EXTRA_TRANSVERSE = 56;
    const ExcelSheet_PAPER_SUPERA = 57;
    const ExcelSheet_PAPER_SUPERB = 58;
    const ExcelSheet_PAPER_US_LETTER_PLUS = 59;
    const ExcelSheet_PAPER_A4_PLUS = 60;
    const ExcelSheet_PAPER_A5_TRANSVERSE = 61;
    const ExcelSheet_PAPER_B5_TRANSVERSE = 62;
    const ExcelSheet_PAPER_A3_EXTRA = 63;
    const ExcelSheet_PAPER_A5_EXTRA = 64;
    const ExcelSheet_PAPER_B5_EXTRA = 65;
    const ExcelSheet_PAPER_A2 = 66;
    const ExcelSheet_PAPER_A3_TRANSVERSE = 67;
    const ExcelSheet_PAPER_A3_EXTRA_TRANSVERSE = 68;
    const ExcelSheet_PAPER_JAPANESE_DOUBLE_POSTCARD = 69;
    const ExcelSheet_PAPER_A6 = 70;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_KAKU2 = 71;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_KAKU3 = 72;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_CHOU3 = 73;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_CHOU4 = 74;
    const ExcelSheet_PAPER_LETTER_ROTATED = 75;
    const ExcelSheet_PAPER_A3_ROTATED = 76;
    const ExcelSheet_PAPER_A4_ROTATED = 77;
    const ExcelSheet_PAPER_A5_ROTATED = 78;
    const ExcelSheet_PAPER_B4_ROTATED = 79;
    const ExcelSheet_PAPER_B5_ROTATED = 80;
    const ExcelSheet_PAPER_JAPANESE_POSTCARD_ROTATED = 81;
    const ExcelSheet_PAPER_DOUBLE_JAPANESE_POSTCARD_ROTATED = 82;
    const ExcelSheet_PAPER_A6_ROTATED = 83;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_KAKU2_ROTATED = 84;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_KAKU3_ROTATED = 85;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_CHOU3_ROTATED = 86;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_CHOU4_ROTATED = 87;
    const ExcelSheet_PAPER_B6 = 88;
    const ExcelSheet_PAPER_B6_ROTATED = 89;
    const ExcelSheet_PAPER_12x11 = 90;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_YOU4 = 91;
    const ExcelSheet_PAPER_JAPANESE_ENVELOPE_YOU4_ROTATED = 92;
    const ExcelSheet_PAPER_PRC16K = 93;
    const ExcelSheet_PAPER_PRC32K = 94;
    const ExcelSheet_PAPER_PRC32K_BIG = 95;
    const ExcelSheet_PAPER_PRC_ENVELOPE1 = 96;
    const ExcelSheet_PAPER_PRC_ENVELOPE2 = 97;
    const ExcelSheet_PAPER_PRC_ENVELOPE3 = 98;
    const ExcelSheet_PAPER_PRC_ENVELOPE4 = 99;
    const ExcelSheet_PAPER_PRC_ENVELOPE5 = 100;
    const ExcelSheet_PAPER_PRC_ENVELOPE6 = 101;
    const ExcelSheet_PAPER_PRC_ENVELOPE7 = 102;
    const ExcelSheet_PAPER_PRC_ENVELOPE8 = 103;
    const ExcelSheet_PAPER_PRC_ENVELOPE9 = 104;
    const ExcelSheet_PAPER_PRC_ENVELOPE10 = 105;
    const ExcelSheet_PAPER_PRC16K_ROTATED = 106;
    const ExcelSheet_PAPER_PRC32K_ROTATED = 107;
    const ExcelSheet_PAPER_PRC32KBIG_ROTATED = 108;
    const ExcelSheet_PAPER_PRC_ENVELOPE1_ROTATED = 109;
    const ExcelSheet_PAPER_PRC_ENVELOPE2_ROTATED = 110;
    const ExcelSheet_PAPER_PRC_ENVELOPE3_ROTATED = 111;
    const ExcelSheet_PAPER_PRC_ENVELOPE4_ROTATED = 112;
    const ExcelSheet_PAPER_PRC_ENVELOPE5_ROTATED = 113;
    const ExcelSheet_PAPER_PRC_ENVELOPE6_ROTATED = 114;
    const ExcelSheet_PAPER_PRC_ENVELOPE7_ROTATED = 115;
    const ExcelSheet_PAPER_PRC_ENVELOPE8_ROTATED = 116;
    const ExcelSheet_PAPER_PRC_ENVELOPE9_ROTATED = 117;
    const ExcelSheet_PAPER_PRC_ENVELOPE10_ROTATED = 118;
    const ExcelSheet_CELLTYPE_EMPTY = 0;
    const ExcelSheet_CELLTYPE_NUMBER = 1;
    const ExcelSheet_CELLTYPE_STRING = 2;
    const ExcelSheet_CELLTYPE_BOOLEAN = 3;
    const ExcelSheet_CELLTYPE_BLANK = 4;
    const ExcelSheet_CELLTYPE_ERROR = 5;
    const ExcelSheet_ERRORTYPE_NULL = 0;
    const ExcelSheet_ERRORTYPE_DIV_0 = 7;
    const ExcelSheet_ERRORTYPE_VALUE = 15;
    const ExcelSheet_ERRORTYPE_REF = 23;
    const ExcelSheet_ERRORTYPE_NAME = 29;
    const ExcelSheet_ERRORTYPE_NUM = 36;
    const ExcelSheet_ERRORTYPE_NA = 42;
    
    /* 
     * Constructor
     * 
     * @param string $sWorkbookFilename File path under which to save the workbook
     * @param string $sLibraryPath Path to the Excel libraries
     */
    public function __construct($sWorkbookFilename, $sLibraryPath)
    {
        self::$__CLASS__ = __CLASS__;
        self::$_INCLUDES = array();
        
        self::$_WORKSHEET_CLASS = 'ExcelSheet';
        self::$_STYLE_CLASS = 'ExcelFormat';
        
        $this->_fDefaultColSize = parent::DEFAULT_COL_SIZE;
        $this->_aColWidth = array();
        
        parent::__construct($sLibraryPath);
        
        $this->_oExcelBook = new \ExcelBook(self::LICENSE_NAME, self::LICENSE_KEY);
        $this->_oExcelBook->setLocale('UTF-8');
        $this->_sWorkbookFilename = $sWorkbookFilename;
        
        $this->_aPictureId = array();
        $this->_aCustomNumberFormat = array();
    }
    
    /*
     * Adds a worksheeet to the Excel workbook
     * 
     * @param string $sWorksheetName Name of the worksheeet to add
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     * @return ExcelSheet Worksheet of type LibXL
     */
    public function addWorksheetToWorkbook($sWorksheetName = '', $iIndexInWorkBook = 0)
    {
        $oWorksheet = $this->_oExcelBook->addSheet($sWorksheetName);
        
        return $oWorksheet;
    }
    
    /*
     * Sets the workhsset as active in the Excel workbook
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     */
    public function setWorksheetAsActiveInWorkbook($oWorksheet, $iIndexInWorkBook)
    {
        $this->_oExcelBook->activeSheet($iIndexInWorkBook);
    }
    
    /*
     * Converts a style to a LibXL one
     * 
     * List of the properties supported during conversion:
     * font         => name
     *              => size
     *              => bold
     *              => italic
     *              => color => argb
     *                       => rgb
     *                       => index
     *              => underline
     * fill         => type
     *              => startcolor => argb
     *                            => rgb
     *                            => index
     * numberformat => code (string format or index)
     * alignment    => horizontal
     *              => vertical
     *              => wrap
     * borders      => allborders => color => argb
     *                                     => rgb
     *                                     => index
     *                            => style
     * 
     * @param array $aStyle Style under array format
     * @return ExcelFormat Style converted to an array that we can add to the LibXL
     * workbbook via the method addFormat()
     */
    protected function _convertAStyleForLibXL($aStyle = array())
    {
        $oConvertedStyle = new \ExcelFormat($this->_oExcelBook);
        
        if (!empty($aStyle)) foreach ($aStyle as $sKey => $value)
        {
            switch($sKey) {
                case 'font':
                    if (!empty($value)) 
                    {
                        $oFont = new \ExcelFont($this->_oExcelBook);
                        foreach ($value as $sKey2 => $value2)
                        {
                            switch($sKey2) {
                                case 'name':
                                    $oFont->name($value2);
                                    break;
                                case 'size':
                                    $oFont->size($value2);
                                    break;
                                case 'bold':
                                    switch($value2) {
                                        case TRUE:
                                            $oFont->bold(TRUE);
                                            break;
                                        case FALSE:
                                            $oFont->bold(FALSE);
                                            break;
                                    }
                                    break;
                                case 'italic':
                                    switch($value2) {
                                        case TRUE:
                                            $oFont->italics(TRUE);
                                            break;
                                        case FALSE:
                                            $oFont->italics(FALSE);
                                            break;
                                    }
                                    break;
                                case 'color':
                                    if (!empty($value2)) foreach ($value2 as $sKey3 => $value3)
                                    {
                                        switch($sKey3) {
                                            case 'argb':
                                                $oFont->color($this->convertARGBToColorIndex($value3));
                                                break;
                                            case 'rgb':
                                                $oFont->color($this->convertRGBToColorIndex($value3));
                                                break;
                                            case 'index':
                                                $oFont->color($value3);
                                                break;
                                        }
                                    }
                                    break;
                                case 'underline':
                                    switch($value2) {
                                        case 'none':
                                            $oFont->underline(self::ExcelFont_UNDERLINE_NONE);
                                            break;
                                        case 'double':
                                            $oFont->underline(self::ExcelFont_UNDERLINE_DOUBLE);
                                            break;
                                        case 'doubleAccounting':
                                            $oFont->underline(self::ExcelFont_UNDERLINE_DOUBLEACC);
                                            break;
                                        case 'single':
                                            $oFont->underline(self::ExcelFont_UNDERLINE_SINGLE);
                                            break;
                                        case 'singleAccounting':
                                            $oFont->underline(self::ExcelFont_UNDERLINE_SINGLEACC);
                                            break;
                                    }
                                    break;
                            }
                        }
                        $this->_oExcelBook->addFont($oFont);
                        $oConvertedStyle->setFont($oFont);
                    }
                    break;
                case 'fill':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {                            
                            case 'type':
                                switch($value2) {
                                    case 'none':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_NONE);
                                        break;
                                    case 'solid':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_SOLID);
                                        break;
                                    case 'linear':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_NONE);
                                        break;
                                    case 'path':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_NONE);
                                        break;
                                    case 'darkDown':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_REVDIAGSTRIPE);
                                        break;
                                    case 'darkGray':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_GRAY75);
                                        break;
                                    case 'darkGrid':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_DIAGCROSSHATCH);
                                        break;
                                    case 'darkHorizontal':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_HORSTRIPE);
                                        break;
                                    case 'darkTrellis':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THICKDIAGCROSSHATCH);
                                        break;
                                    case 'darkUp':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_DIAGSTRIPE);
                                        break;
                                    case 'darkVertical':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_VERSTRIPE);
                                        break;
                                    case 'gray0625':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_GRAY6P25);
                                        break;
                                    case 'gray125':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_GRAY12P5);
                                        break;
                                    case 'lightDown':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINREVDIAGSTRIPE);
                                        break;
                                    case 'lightGray':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_GRAY25);
                                        break;
                                    case 'lightGrid':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINHORCROSSHATCH);
                                        break;
                                    case 'lightHorizontal':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINHORSTRIPE);
                                        break;
                                    case 'lightTrellis':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINDIAGCROSSHATCH);
                                        break;
                                    case 'lightUp':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINDIAGSTRIPE);
                                        break;
                                    case 'lightVertical':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_THINVERSTRIPE);
                                        break;
                                    case 'mediumGray':
                                        $oConvertedStyle->fillPattern(self::ExcelFormat_FILLPATTERN_GRAY50);
                                        break;
                                }
                                break;
                            case 'startcolor':
                                if (!empty($value2)) foreach ($value2 as $sKey3 => $value3)
                                {
                                    switch($sKey3) {
                                        case 'argb':
                                            $oConvertedStyle->patternForegroundColor($this->convertARGBToColorIndex($value3));
                                            break;
                                        case 'rgb':
                                            $oConvertedStyle->patternForegroundColor($this->convertRGBToColorIndex($value3));
                                            break;
                                        case 'index':
                                            $oConvertedStyle->patternForegroundColor($value3);
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                    break;
                case 'numberformat':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {
                            case 'code':
                                $oConvertedStyle->numberFormat($this->_convertNumFormatCodeToIndex($value2));
                                break;
                        }
                    }
                    break;
                
                case 'alignment':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {
                            case 'horizontal':
                                switch($value2) {
                                    case 'general':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_GENERAL);
                                        break;
                                    case 'left':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_LEFT);
                                        break;
                                    case 'right':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_RIGHT);
                                        break;
                                    case 'center':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_CENTER);
                                        break;
                                    case 'centerContinuous':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_MERGE);
                                        break;
                                    case 'justify':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_JUSTIFY);
                                        break;
                                    case 'distributed':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_DISTRIBUTED);
                                        break;
                                    case 'fill':
                                        $oConvertedStyle->horizontalAlign(self::ExcelFormat_ALIGNH_FILL);
                                        break;
                                }
                                break;
                            case 'vertical':
                                switch($value2) {
                                    case 'bottom':
                                        $oConvertedStyle->verticalAlign(self::ExcelFormat_ALIGNV_BOTTOM);
                                        break;
                                    case 'top':
                                        $oConvertedStyle->verticalAlign(self::ExcelFormat_ALIGNV_TOP);
                                        break;
                                    case 'center':
                                        $oConvertedStyle->verticalAlign(self::ExcelFormat_ALIGNV_CENTER);
                                        break;
                                    case 'justify':
                                        $oConvertedStyle->verticalAlign(self::ExcelFormat_ALIGNV_JUSTIFY);
                                        break;
                                    case 'distributed':
                                        $oConvertedStyle->verticalAlign(self::ExcelFormat_ALIGNV_DISTRIBUTED);
                                        break;
                                }
                                break;
                            case 'wrap':
                                switch($value2) {
                                    case TRUE:
                                        $oConvertedStyle->wrap(TRUE);
                                        break;
                                    case FALSE:
                                        $oConvertedStyle->wrap(FALSE);
                                        break;
                                }
                                break;
                        }
                    }
                    break;
                
                case 'borders':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {
                            case 'allborders':
                                if (!empty($value2)) foreach ($value2 as $sKey3 => $value3)
                                {
                                    switch($sKey3) {
                                        case 'color':
                                            if (!empty($value3)) foreach ($value3 as $sKey4 => $value4)
                                            {
                                                switch($sKey4) {
                                                    case 'argb':
                                                        $oConvertedStyle->borderColor($this->convertARGBToColorIndex($value4));
                                                        break;
                                                    case 'rgb':
                                                        $oConvertedStyle->borderColor($this->convertRGBToColorIndex($value4));
                                                        break;
                                                    case 'index':
                                                        $oConvertedStyle->borderColor($value4);
                                                        break;
                                                }
                                            }
                                            break;                                    
                                        case 'style':
                                            switch($value3) {
                                                case 'none':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_NONE);
                                                    break;
                                                case 'dashDot':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_DASHDOT);
                                                    break;
                                                case 'dashDotDot':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_ORDERSTYLE_DASHDOTDOT);
                                                    break;
                                                case 'dashed':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_DASHED);
                                                    break;
                                                case 'dotted':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_DOTTED);
                                                    break;
                                                case 'double':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_DOUBLE);
                                                    break;
                                                case 'hair':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_HAIR);
                                                    break;
                                                case 'medium':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_MEDIUM);
                                                    break;
                                                case 'mediumDashDot':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_MEDIUMDASHDOT);
                                                    break;
                                                case 'mediumDashDotDot':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_MEDIUMDASHDOTDOT);
                                                    break;
                                                case 'mediumDashed':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_MEDIUMDASHED);
                                                    break;
                                                case 'slantDashDot':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_SLANTDASHDOT);
                                                    break;
                                                case 'thick':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_THICK);
                                                    break;
                                                case 'thin':
                                                    $oConvertedStyle->borderStyle(self::ExcelFormat_BORDERSTYLE_THIN);
                                                    break;
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                    break;
            }
        }   
        
        return $oConvertedStyle;
    }
    
    /*
     * Adds a style to the Excel workbook
     * 
     * @param $aStyle Style under array format
     * @return ExcelFormat Style of type LibXL
     */
    public function addStyleToWorkbook(&$aStyle = array())
    {
        $oStyle = $this->_oExcelBook->addFormat($this->_convertAStyleForLibXL($aStyle));
        
        return $oStyle;
    }
    
    /*
     * Closes the Excel workbook saving it
     */
    public function closeWorkbook()
    {
        //Application of the columnn widths before writing the workbook because
        //they can be define only one time for LibXL
        if (!empty($this->_aColWidth)) foreach($this->_aColWidth as $iIndexInWorkbook => $aColWidth)
        {
            if (!empty($aColWidth)) foreach($aColWidth as $iCol => $fColWidth)
                $this->_oExcelBook->getSheet($iIndexInWorkbook)->setColWidth($iCol, $iCol, $fColWidth);
        }
        
        $this->_oExcelBook->save($this->_sWorkbookFilename);
    }
    
    /*
     * Sets the width of one or many columns of an Excel worksheet
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param float $fWidth Width
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstcol Index of the first column
     * @param int $iLastcol Index of the last column
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function setColumnsWidthToWorksheet($oWorksheet, $fWidth, $iIndexInWorkbook = 0, $iFirstcol = 0, $iLastcol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            for ($iCol = $iFirstcol; $iCol <= $iLastcol; $iCol++) {
                if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                    $this->_aColWidth[$iIndexInWorkbook] = array();
                
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $fWidth;
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Insertion of an image into an Excel worksheet. Supports the types .bmp, 
     * .png, .jpg
     * 
     * @param ExcelSheet $oWorksheet Workhsheet of type LibXL
     * @param string $sFilename File path of the image
     * @param int $iRow Index of the cell line where is inserted the image
     * @param int $iCol Index of the cell column where is inserted the image
     * @return ExcelSheet|bool Workhsheet of type LibXL or false
     */
    public function insertImageToWorksheet($oWorksheet, $sFilename, $iRow = 0, $iCol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (!in_array($sFilename, $this->_aPictureId))
            {
                $iPictureId = $this->_oExcelBook->addPictureFromFile($sFilename);
                $this->_aPictureId[$iPictureId] = $sFilename;
            }
            else
                $iPictureId = array_search($sFilename, $this->_aPictureId);
            
            $aDimensions = getimagesize($sFilename);
            
            $oWorksheet->addPictureDim($iRow, $iCol, $iPictureId, $aDimensions[0], $aDimensions[1]);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Permits to calculate and specify the width of a column after its content
     * 
     * @param ExcelSheet $oWorksheet Workhsset of type LibXL
     * @param int $iIndexInWorkbook Index of the workbook worksheet in façade
     * @param int $iCol Index of the column
     * @param string $pCellToken Cell content of type NULL, string, int, float, double, bool
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param string $sCellContentType Type of cell content eg. 'rowhead'
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Excact calculation of the column 
     * widths after the TrueType fonts
     */
    protected function _setColumnAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth && !empty($aStyle))
        {
            if ($iNbCharPercent > -1)
            {
                $sPercentToken = str_pad(',00 %', mb_strlen((string)(int)floor($pCellToken), 'UTF-8')+5, '0', STR_PAD_LEFT);
                $fCalculatedWidth = $this->calculateExactColumnWidth($aStyle['font']['name'], $aStyle['font']['bold'], $aStyle['font']['italic'], $aStyle['font']['size'], $sPercentToken, 0.0, self::DEFAULT_FONT_NAME,  self::DEFAULT_FONT_SIZE);
            }
            else
                $fCalculatedWidth = $this->calculateExactColumnWidth($aStyle['font']['name'], $aStyle['font']['bold'], $aStyle['font']['italic'], $aStyle['font']['size'], (string)$pCellToken, 0.0, self::DEFAULT_FONT_NAME,  self::DEFAULT_FONT_SIZE);
            
            $fCalculatedWidth++;
        }
        else
        {
            if ($iNbCharPercent > -1)
                $fCalculatedWidth = $iNbCharPercent*1.26;
            else
                $fCalculatedWidth = mb_strlen((string)$pCellToken, 'UTF-8')*1.26;
        }
        $fCurrentWidth = $this->_aColWidth[$iIndexInWorkbook][$iCol];
        
        if ($fCalculatedWidth > $fCurrentWidth)
            $this->_aColWidth[$iIndexInWorkbook][$iCol] = $fCalculatedWidth;
    }
    
    /*
     * Permits to calculate and specify the width of the columns of cells merged 
     * horizontally depending on the content of the upper left cell
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstCol Index of the first column of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of 
     * the columns according to TrueType fonts
     */
    protected function _setMergedColumnsAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstCol = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth && !empty($aStyle))
        {
            if ($iNbCharPercent > -1)
            {
                $sPercentToken = str_pad(',00 %', mb_strlen((string)(int)floor($pCellToken), 'UTF-8')+5, '0', STR_PAD_LEFT);
                $fCalculatedWidth = $this->calculateExactColumnWidth($aStyle['font']['name'], $aStyle['font']['bold'], $aStyle['font']['italic'], $aStyle['font']['size'], $sPercentToken, 0.0, self::DEFAULT_FONT_NAME, self::DEFAULT_FONT_SIZE);
            }
            else
                $fCalculatedWidth = $this->calculateExactColumnWidth($aStyle['font']['name'], $aStyle['font']['bold'], $aStyle['font']['italic'], $aStyle['font']['size'], (string)$pCellToken, 0.0, self::DEFAULT_FONT_NAME, self::DEFAULT_FONT_SIZE);
            
            $fCalculatedWidth++;
        }
        else
        {
            if ($iNbCharPercent > -1)
                $fCalculatedWidth = $iNbCharPercent*1.26;
            else
                $fCalculatedWidth = mb_strlen((string)$pCellToken, 'UTF-8')*1.26;
        }
        
        $iNbCol = $iLastCol - $iFirstCol + 1;
        $fSumColWidth = 0.0;
        
        for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
            $fSumColWidth += $this->_aColWidth[$iIndexInWorkbook][$iCol];
        }
        
        if ($fCalculatedWidth > $fSumColWidth)
        {
            $fRestWidthPerCol = ($fCalculatedWidth - $fSumColWidth) / $iNbCol;
            
            for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                $fColWidth = $this->_aColWidth[$iIndexInWorkbook][$iCol];
                
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $fColWidth + $fRestWidthPerCol;
            }
        }
    }
    
    /*
     * Writes in a range of cells with a unique style. The style isn't overwritten
     * if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line to write
     * @param int $iFirstCol Index of the first column to write
     * @param int $iLastRow Index of the last line of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param array $aCellMatrix Cell matrix of the form $aCellMatrix[$iRow][$iCol] = $pValue
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param array $aMaxStringWithLengthPerCol Array indexed by column giving
     * the field of the column of maximal size with its length
     * @param array $aIndexColCellsWithString Index of the cells containing a
     * string
     * @return ExcelSheet|bool $oWorksheet Worksheet of type LibXL or false
     */
    public function writeFromArrayToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), $oStyle = NULL, &$aStyle = array(), &$aMaxStringWithLengthPerCol = array(), &$aIndexColCellsWithString = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (!empty($aCellMatrix))
            {
                if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
                
                foreach ($aCellMatrix as $aCol) {
                    if (!empty($aCol))
                    {
                        $oWorksheet->writeCol($iFirstCol, $aCol, $iFirstRow, $oStyle);
                    }
                        
                    ++$iFirstCol;
                }
                
                //Initialization of the size of the columns
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++)
                {
                    if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                        $this->_aColWidth[$iIndexInWorkbook] = array();
                    if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
                    {
                        $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
                    }
                }
                
                if (!empty($aMaxStringWithLengthPerCol)) foreach ($aMaxStringWithLengthPerCol as $iCol => $aStringLength)
                {
                    if (!empty($aStringLength) && !empty($aCellMatrix[$iCol])) $this->_setColumnAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook, $iCol, $aStringLength[2], $oStyle, $aStyle, $aStringLength[1]);
                }
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes with a style into a cell of an Excel worksheet. The style isn't
     * overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the cell line where is written
     * @param int $iCol Index of the cell column where is written
     * @param $pToken Content to write of type NULL, string, int, float, double, bool
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param string $sCellContentType Type of cell content eg. 'rowhead'
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function writeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pToken = NULL, $oStyle = NULL, &$aStyle = array(), $sCellContentType = '', $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write($iRow, $iCol, $pToken, $oStyle);
            
            //Initialisation de la taille de la colonne
            if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                $this->_aColWidth[$iIndexInWorkbook] = array();
            if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
            {
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
            }
            
            if ('infohead' != $sCellContentType && !$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook, $iCol, $pToken, $oStyle, $aStyle);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a string with a style into a cell of an Excel worksheet.
     * The style isn't overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iRow Index of the cell line where is written
     * @param int $iCol Index of the cell column where is written
     * @param string $sValue Content to write
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function writeStringToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $sValue = '', $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write($iRow, $iCol, $sValue, $oStyle);
            
            //Initialisation de la taille de la colonne
            if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                $this->_aColWidth[$iIndexInWorkbook] = array();
            if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
            {
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
            }
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook, $iCol, $sValue, $oStyle, $aStyle);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a number with a style into a cell of an Excel worksheet. The style 
     * isn't overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iRow Index of the cell line where is written
     * @param int $iCol Index of the cell column where is written
     * @param $pNumber Content to write of type string, int, float, double
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function writeNumberToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pNumber = 0, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write($iRow, $iCol, $pNumber, $oStyle, self::ExcelFormat_AS_NUMERIC_STRING);
            
            //Initialization of the size of the column
            if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                $this->_aColWidth[$iIndexInWorkbook] = array();
            if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
            {
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
            }
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook, $iCol, $pNumber, $oStyle, $aStyle, $iNbCharPercent);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes an empty cell with a style into an Excel worksheet. The style 
     * isn't overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iRow Index of the cell line
     * @param int $iCol Index of the cell column
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function writeBlankToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write($iRow, $iCol, NULL, $oStyle);
            
            //Initialisation de la taille de la colonne
            if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                $this->_aColWidth[$iIndexInWorkbook] = array();
            if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
            {
                $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes empty cells with a style into an Excel worksheet. The style 
     * isn't overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function writeBlankToManyCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                    $oWorksheet->write($iRow, $iCol, NULL, $oStyle);
                    
                    //Initialisation de la taille de la colonne
                    if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                        $this->_aColWidth[$iIndexInWorkbook] = array();
                    if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
                    {
                        $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
                    }
                }
            }
                    
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Apply a style to cells within an Excel worksheet. The style 
     * isn't overwritten if NULL.
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $indexInWorkbook Index of workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function setStyleToManyCellsToWorksheet($oWorksheet, $indexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            if (isset($oStyle))
            {
                for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                    for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                        $oWorksheet->setCellFormat($iRow, $iCol, $oStyle);
                    }
                }
            }
                    
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Merge cells in an Excel worksheet
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * 
     * Only to resize the columns:
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * 
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function mergeCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->setMerge($iFirstRow, $iLastRow, $iFirstCol, $iLastCol);
            
            //Initialisation de la taille des colonnes
            for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                if (!isset($this->_aColWidth[$iIndexInWorkbook]))
                    $this->_aColWidth[$iIndexInWorkbook] = array();
                if (!isset($this->_aColWidth[$iIndexInWorkbook][$iCol]))
                {
                    $this->_aColWidth[$iIndexInWorkbook][$iCol] = $this->_fDefaultColSize;
                }
            }
            
            if (!is_null($pCellToken) && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setMergedColumnsAutosizeToWorksheet($oWorksheet, $iIndexInWorkbook, $iFirstCol, $iLastCol, $pCellToken, $oStyle, $aStyle, $iNbCharPercent);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Deletes a column of an Excel worksheet
     * 
     * @param ExcelSheet $oWorksheet Worksheet of type LibXL
     * @param int $iIndexInWorkbook Index of workbook worksheet in facade
     * @param int $iCol Index of the column to delete
     * @return ExcelSheet|bool Worksheet of type LibXL or false
     */
    public function deleteColumnInWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iCol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->removeCol($iCol, $iCol);
            
            //Décalage des indices pour les largeurs de colonnes
            $aNewColWidth = array();
            
            if (!empty($this->_aColWidth)) foreach($this->_aColWidth as $iTmpIndexInWorkbook => $aColWidth)
            {
                if ($iTmpIndexInWorkbook == $iIndexInWorkbook)
                {
                    if (!isset($aNewColWidth[$iTmpIndexInWorkbook])) $aNewColWidth[$iTmpIndexInWorkbook] = array();
                    
                    if (!empty($aColWidth)) foreach($aColWidth as $iTmpCol => $fColWidth)
                    {
                        if ($iTmpCol < $iCol)
                            $aNewColWidth[$iTmpIndexInWorkbook][$iTmpCol] = $fColWidth;
                        elseif ($iTmpCol > $iCol)
                            $aNewColWidth[$iTmpIndexInWorkbook][$iTmpCol-1] = $fColWidth;
                    }
                }
                else
                    $aNewColWidth[$iTmpIndexInWorkbook] = $this->_aColWidth[$iTmpIndexInWorkbook];
            }
            
            unset($this->_aColWidth);
            $this->_aColWidth = &$aNewColWidth;
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Converts a format of an Excel cell into its integer index creating a new
     * if needed
     * 
     * @param $pNumFormat Format code of type string or entier for index
     * @return int Index of format code
     */
    protected function _convertNumFormatCodeToIndex($pNumFormat = '')
    {
        if (is_string($pNumFormat))
        {
            switch (true) {
                case $pNumFormat === 'General':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_GENERAL;
                    break;
                case $pNumFormat === '0':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER;
                    break;
                case $pNumFormat === '0.00':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_D2;
                    break;
                case $pNumFormat === '#,##0':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_SEP;
                    break;
                case $pNumFormat === '#,##0.00':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_SEP_D2;
                    break;
                case $pNumFormat === '0$;(0$)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CURRENCY_NEGBRA;
                    break;
                case $pNumFormat === '0$;[Red](0$)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CURRENCY_NEGBRARED;
                    break;
                case $pNumFormat === '$0.00;($0.00)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CURRENCY_D2_NEGBRA;
                    break;
                case $pNumFormat === '$0.00;[Red]($0.00)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CURRENCY_D2_NEGBRARED;
                    break;
                case $pNumFormat === '0%':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_PERCENT;
                    break;
                case $pNumFormat === '0.00%':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_PERCENT_D2;
                    break;
                case $pNumFormat === '0.00E+00':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_SCIENTIFIC_D2;
                    break;
                case $pNumFormat === '#" "?/?':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_FRACTION_ONEDIG;
                    break;
                case $pNumFormat === '#" "??/??':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_FRACTION_TWODIG;
                    break;
                case $pNumFormat === 'mm-dd-yy':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_DATE;
                    break;
                case $pNumFormat === 'd-mmm-yy':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_D_MON_YY;
                    break;
                case $pNumFormat === 'd-mmm':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_D_MON;
                    break;
                case $pNumFormat === 'mmm-yy':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_MON_YY;
                    break;
                case $pNumFormat === 'h:mm AM/PM':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_HMM_AM;
                    break;
                case $pNumFormat === 'h:mm:ss AM/PM':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_HMMSS_AM;
                    break;
                case $pNumFormat === 'h:mm':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_HMM;
                    break;
                case $pNumFormat === 'h:mm:ss':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_HMMSS;
                    break;
                case $pNumFormat === 'm/d/yy h:mm':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_MDYYYY_HMM;
                    break;
                case $pNumFormat === '#,##0;(#,##0)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_SEP_NEGBRA;
                    break;
                case $pNumFormat === '#,##0;[Red](#,##0)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_SEP_NEGBRARED;
                    break;
                case $pNumFormat === '#,##0.00;(#,##0.00)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_D2_SEP_NEGBRA;
                    break;
                case $pNumFormat === '#,##0.00;[Red](#,##0.00)':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_NUMBER_D2_SEP_NEGBRARED;
                    break;
                case $pNumFormat === '_-#,##0_-;_-"-"#,##0_-;_-@_-':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_ACCOUNT;
                    break;
                case $pNumFormat === '_-$* #,##0_-;_-$* "-"#,##0_-;_-$* "-"_-;_-@_-':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_ACCOUNTCUR;
                    break;
                case $pNumFormat === '_-#,##0.00_-;_-"-"#,##0.00_-;_-"-"??_-;_-@_-':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_ACCOUNT_D2;
                    break;
                case $pNumFormat === '_-$* #,##0.00_-;_-$* "-"#,##0.00_-;_-$* "-"??_-;_-@_-':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_ACCOUNT_D2_CUR;
                    break;
                case $pNumFormat === 'mm:ss':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_MMSS;
                    break;
                case $pNumFormat === '[h]:mm:ss':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_H0MMSS;
                    break;
                case $pNumFormat === 'mmss.0':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_MMSS0;
                    break;
                case $pNumFormat === '##0.0E+0':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_CUSTOM_000P0E_PLUS0;
                    break;
                case $pNumFormat === '@':
                    $iNumberFormat = self::ExcelFormat_NUMFORMAT_TEXT;
                    break;
                default:
                    if (!in_array($pNumFormat, $this->_aCustomNumberFormat)) {
                        $iNumberFormat = $this->_oExcelBook->addCustomFormat($pNumFormat);
                        $this->_aCustomNumberFormat[$iNumberFormat] = $pNumFormat;
                    }
                    else
                        $iNumberFormat = array_search($pNumFormat, $this->_aCustomNumberFormat);
                    break;
            }
        }
        else
            $iNumberFormat = (int)$pNumFormat;
            
        return $iNumberFormat;
    }
    
    /*
     * Sets a number format in an Excel style
     * 
     * @param ExcelFormat $oStyle Style of type LibXL
     * @param $pNumFormat Number format of type string or entier for index
     * @return ExcelFormat|bool Style of type LibXL or false
     */
    public function setNumFormatToStyle($oStyle, $pNumFormat = '')
    {
        if ($this->checkStyleClass($oStyle))
        {
            $oStyle->numberFormat($this->_convertNumFormatCodeToIndex($pNumFormat));
        
            return $oStyle;
        }
        else
            return FALSE;
    }
    
    /*
     * Specifies the text wrap in an Excel style
     * 
     * @param ExcelFormat $oStyle Style de type LibXL
     * @return ExcelFormat|bool Style of type LibXL or false
     */
    public function setTextWrapToStyle($oStyle)
    {
        if ($this->checkStyleClass($oStyle))
        {
            $oStyle->wrap(TRUE);
            
            return $oStyle;
        }
        else
            return FALSE;
    }
}