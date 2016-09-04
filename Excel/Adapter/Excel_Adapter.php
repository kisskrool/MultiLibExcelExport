<?php

namespace MultiLibExcelExport\Excel\Adapter;

class Excel_Adapter
{
    static protected $__CLASS__;
    static protected $_LIBRARY_PATH;
    static protected $_INCLUDES;
    static protected $_WORKSHEET_CLASS;
    static protected $_STYLE_CLASS;
    
    const PATH_TRUETYPE_FONTS = '/Excel_Fonts/';
    const DEFAULT_COL_SIZE = 8.43;//10.71 by default under Excel
    const DEFAULT_ROW_HEIGHT = 12.75;//12.75 by default under Excel
    
    const DEFAULT_FONT_NAME = 'Arial';
    const DEFAULT_FONT_SIZE = 10;
    
    /*
     * TTF font files
     */
    const FONT_FILENAME_ARIAL = 'arial.ttf';
    const FONT_FILENAME_ARIAL_BOLD = 'arialbd.ttf';
    const FONT_FILENAME_ARIAL_ITALIC = 'ariali.ttf';
    const FONT_FILENAME_ARIAL_BOLD_ITALIC = 'arialbi.ttf';
    const FONT_FILENAME_CALIBRI = 'CALIBRI.TTF';
    const FONT_FILENAME_CALIBRI_BOLD = 'CALIBRIB.TTF';
    const FONT_FILENAME_CALIBRI_ITALIC = 'CALIBRII.TTF';
    const FONT_FILENAME_CALIBRI_BOLD_ITALIC = 'CALIBRIZ.TTF';
    const FONT_FILENAME_COMIC_SANS_MS = 'comic.ttf';
    const FONT_FILENAME_COMIC_SANS_MS_BOLD = 'comicbd.ttf';
    const FONT_FILENAME_COURIER_NEW = 'cour.ttf';
    const FONT_FILENAME_COURIER_NEW_BOLD = 'courbd.ttf';
    const FONT_FILENAME_COURIER_NEW_ITALIC = 'couri.ttf';
    const FONT_FILENAME_COURIER_NEW_BOLD_ITALIC = 'courbi.ttf';
    const FONT_FILENAME_GEORGIA = 'georgia.ttf';
    const FONT_FILENAME_GEORGIA_BOLD = 'georgiab.ttf';
    const FONT_FILENAME_GEORGIA_ITALIC = 'georgiai.ttf';
    const FONT_FILENAME_GEORGIA_BOLD_ITALIC = 'georgiaz.ttf';
    const FONT_FILENAME_IMPACT = 'impact.ttf';
    const FONT_FILENAME_LIBERATION_SANS = 'LiberationSans-Regular.ttf';
    const FONT_FILENAME_LIBERATION_SANS_BOLD = 'LiberationSans-Bold.ttf';
    const FONT_FILENAME_LIBERATION_SANS_ITALIC = 'LiberationSans-Italic.ttf';
    const FONT_FILENAME_LIBERATION_SANS_BOLD_ITALIC = 'LiberationSans-BoldItalic.ttf';
    const FONT_FILENAME_LUCIDA_CONSOLE = 'lucon.ttf';
    const FONT_FILENAME_LUCIDA_SANS_UNICODE = 'l_10646.ttf';
    const FONT_FILENAME_MICROSOFT_SANS_SERIF = 'micross.ttf';
    const FONT_FILENAME_PALATINO_LINOTYPE = 'pala.ttf';
    const FONT_FILENAME_PALATINO_LINOTYPE_BOLD = 'palab.ttf';
    const FONT_FILENAME_PALATINO_LINOTYPE_ITALIC = 'palai.ttf';
    const FONT_FILENAME_PALATINO_LINOTYPE_BOLD_ITALIC = 'palabi.ttf';
    const FONT_FILENAME_SYMBOL = 'symbol.ttf';
    const FONT_FILENAME_TAHOMA = 'tahoma.ttf';
    const FONT_FILENAME_TAHOMA_BOLD = 'tahomabd.ttf';
    const FONT_FILENAME_TIMES_NEW_ROMAN = 'times.ttf';
    const FONT_FILENAME_TIMES_NEW_ROMAN_BOLD = 'timesbd.ttf';
    const FONT_FILENAME_TIMES_NEW_ROMAN_ITALIC = 'timesi.ttf';
    const FONT_FILENAME_TIMES_NEW_ROMAN_BOLD_ITALIC = 'timesbi.ttf';
    const FONT_FILENAME_TREBUCHET_MS = 'trebuc.ttf';
    const FONT_FILENAME_TREBUCHET_MS_BOLD = 'trebucbd.ttf';
    const FONT_FILENAME_TREBUCHET_MS_ITALIC = 'trebucit.ttf';
    const FONT_FILENAME_TREBUCHET_MS_BOLD_ITALIC = 'trebucbi.ttf';
    const FONT_FILENAME_VERDANA = 'verdana.ttf';
    const FONT_FILENAME_VERDANA_BOLD = 'verdanab.ttf';
    const FONT_FILENAME_VERDANA_ITALIC = 'verdanai.ttf';
    const FONT_FILENAME_VERDANA_BOLD_ITALIC = 'verdanaz.ttf';
    
    /*
     * Width of column by default for a font of a given size. Empirical data
     * under Excel 2007.
     */
    static protected $_XL2007_DEFAULT_FONT_COLUMN_WIDTHS;
    
    /*
     * Palette by default of Excel 97 in format
     * colorIndex => array(R, G, B, A)
     */
    static protected $_XL97_COLOR_PALETTE;
    
    /*
     * Constructor making inclusion of the files of the Excel library
     * 
     * @param string Path of the library
     */
    public function __construct($sLibraryPath)
    {
        self::$__CLASS__ = __CLASS__;
        //The initialization of static variables is made here in order to be compatible
        //with parralelization using the extension pthreads v2
        static::$_LIBRARY_PATH = $sLibraryPath;
        
        static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS = array(
            'Arial' => array(
                1 => array('px' => 24, 'width' => 12.00000000),
                2 => array('px' => 24, 'width' => 12.00000000),
                3 => array('px' => 32, 'width' => 10.66406250),
                4 => array('px' => 32, 'width' => 10.66406250),
                5 => array('px' => 40, 'width' => 10.00000000),
                6 => array('px' => 48, 'width' => 9.59765625),
                7 => array('px' => 48, 'width' => 9.59765625),
                8 => array('px' => 56, 'width' => 9.33203125),
                9 => array('px' => 64, 'width' => 9.14062500),
                10 => array('px' => 64, 'width' => 9.14062500),
            ),
            'Calibri' => array(
                1 => array('px' => 24, 'width' => 12.00000000),
                2 => array('px' => 24, 'width' => 12.00000000),
                3 => array('px' => 32, 'width' => 10.66406250),
                4 => array('px' => 32, 'width' => 10.66406250),
                5 => array('px' => 40, 'width' => 10.00000000),
                6 => array('px' => 48, 'width' => 9.59765625),
                7 => array('px' => 48, 'width' => 9.59765625),
                8 => array('px' => 56, 'width' => 9.33203125),
                9 => array('px' => 56, 'width' => 9.33203125),
                10 => array('px' => 64, 'width' => 9.14062500),
                11 => array('px' => 64, 'width' => 9.14062500),
            ),
            'Verdana' => array(
                1 => array('px' => 24, 'width' => 12.00000000),
                2 => array('px' => 24, 'width' => 12.00000000),
                3 => array('px' => 32, 'width' => 10.66406250),
                4 => array('px' => 32, 'width' => 10.66406250),
                5 => array('px' => 40, 'width' => 10.00000000),
                6 => array('px' => 48, 'width' => 9.59765625),
                7 => array('px' => 48, 'width' => 9.59765625),
                8 => array('px' => 64, 'width' => 9.14062500),
                9 => array('px' => 72, 'width' => 9.00000000),
                10 => array('px' => 72, 'width' => 9.00000000),
            ),
        );
        
        static::$_XL97_COLOR_PALETTE = array(
            0x08 => array(0x00, 0x00, 0x00, 0x00), //black
            0x09 => array(0xff, 0xff, 0xff, 0x00), //white
            0x0A => array(0xff, 0x00, 0x00, 0x00), //red
            0x0B => array(0x00, 0xff, 0x00, 0x00), //lime
            0x0C => array(0x00, 0x00, 0xff, 0x00), //blue
            0x0D => array(0xff, 0xff, 0x00, 0x00), //yellow
            0x0E => array(0xff, 0x00, 0xff, 0x00), //magenta
            0x0F => array(0x00, 0xff, 0xff, 0x00), //cyan
            0x10 => array(0x80, 0x00, 0x00, 0x00), //brown
            0x11 => array(0x00, 0x80, 0x00, 0x00), //green
            0x12 => array(0x00, 0x00, 0x80, 0x00), //navy
            0x13 => array(0x80, 0x80, 0x00, 0x00),
            0x14 => array(0x80, 0x00, 0x80, 0x00), //purple
            0x15 => array(0x00, 0x80, 0x80, 0x00),
            0x16 => array(0xc0, 0xc0, 0xc0, 0x00), //silver
            0x17 => array(0x80, 0x80, 0x80, 0x00), //gray
            0x18 => array(0x99, 0x99, 0xff, 0x00),
            0x19 => array(0x99, 0x33, 0x66, 0x00),
            0x1A => array(0xff, 0xff, 0xcc, 0x00),
            0x1B => array(0xcc, 0xff, 0xff, 0x00),
            0x1C => array(0x66, 0x00, 0x66, 0x00),
            0x1D => array(0xff, 0x80, 0x80, 0x00),
            0x1E => array(0x00, 0x66, 0xcc, 0x00),
            0x1F => array(0xcc, 0xcc, 0xff, 0x00),
            0x20 => array(0x00, 0x00, 0x80, 0x00),
            0x21 => array(0xff, 0x00, 0xff, 0x00),
            0x22 => array(0xff, 0xff, 0x00, 0x00),
            0x23 => array(0x00, 0xff, 0xff, 0x00),
            0x24 => array(0x80, 0x00, 0x80, 0x00),
            0x25 => array(0x80, 0x00, 0x00, 0x00),
            0x26 => array(0x00, 0x80, 0x80, 0x00),
            0x27 => array(0x00, 0x00, 0xff, 0x00),
            0x28 => array(0x00, 0xcc, 0xff, 0x00),
            0x29 => array(0xcc, 0xff, 0xff, 0x00),
            0x2A => array(0xcc, 0xff, 0xcc, 0x00),
            0x2B => array(0xff, 0xff, 0x99, 0x00),
            0x2C => array(0x99, 0xcc, 0xff, 0x00),
            0x2D => array(0xff, 0x99, 0xcc, 0x00),
            0x2E => array(0xcc, 0x99, 0xff, 0x00),
            0x2F => array(0xff, 0xcc, 0x99, 0x00),
            0x30 => array(0x33, 0x66, 0xff, 0x00),
            0x31 => array(0x33, 0xcc, 0xcc, 0x00),
            0x32 => array(0x99, 0xcc, 0x00, 0x00),
            0x33 => array(0xff, 0xcc, 0x00, 0x00),
            0x34 => array(0xff, 0x99, 0x00, 0x00),
            0x35 => array(0xff, 0x66, 0x00, 0x00), //orange
            0x36 => array(0x66, 0x66, 0x99, 0x00),
            0x37 => array(0x96, 0x96, 0x96, 0x00),
            0x38 => array(0x00, 0x33, 0x66, 0x00),
            0x39 => array(0x33, 0x99, 0x66, 0x00),
            0x3A => array(0x00, 0x33, 0x00, 0x00),
            0x3B => array(0x33, 0x33, 0x00, 0x00),
            0x3C => array(0x99, 0x33, 0x00, 0x00),
            0x3D => array(0x99, 0x33, 0x66, 0x00),
            0x3E => array(0x33, 0x33, 0x99, 0x00),
            0x3F => array(0x33, 0x33, 0x33, 0x00),
        );
        
        if (!empty(static::$_INCLUDES)) foreach(static::$_INCLUDES as $sInclude)
        {
            require_once static::$_LIBRARY_PATH . $sInclude;
        }
    }
    
    /*
     * Converts a color in ARGB format to index of Excel 97 palette if it exits
     * 
     * @param string $sARGB eg. '00FF00AA'
     * @return int Color index
     */
    public function convertARGBToColorIndex($sARGB = '')
    {
        if ($this->checkLengthString($sARGB, 8))
        {
            $a = 0x00;//always 0 for WriteExcel
            $r = hexdec(substr($sARGB, 2, 2));
            $g = hexdec(substr($sARGB, 4, 2));
            $b = hexdec(substr($sARGB, 6, 2));

            return $this->_convertRGBToColorIndex($r, $g, $b, $a);
        }
        else
            return FALSE;
    }
    
    /*
     * Converts an index Excel 97 palette to an ARGB color if it exists
     * 
     * @param int $iIndex Color index
     * @return string $sARGB eg. '00FF00AA'
     */
    public function convertColorIndexToARGB($iIndex)
    {
        if (isset(static::$_XL97_COLOR_PALETTE[$iIndex]))
        {
            $sARGB = str_pad(dechex(static::$_XL97_COLOR_PALETTE[$iIndex][3]), 2, '0', STR_PAD_LEFT);//a
            $sARGB .= str_pad(dechex(static::$_XL97_COLOR_PALETTE[$iIndex][0]), 2, '0', STR_PAD_LEFT);//r
            $sARGB .= str_pad(dechex(static::$_XL97_COLOR_PALETTE[$iIndex][1]), 2, '0', STR_PAD_LEFT);//g
            $sARGB .= str_pad(dechex(static::$_XL97_COLOR_PALETTE[$iIndex][2]), 2, '0', STR_PAD_LEFT);//b
            
            return strtoupper($sARGB);
        }
        else
            return FALSE;
    }
    
    /*
     * Converts a color from RGB format to index of Excel 97 palette if it exists 
     * 
     * @param string $sRGB ex. 'FF00AA'
     * @return int Index couleur
     */
    public function convertRGBToColorIndex($sRGB = '')
    {
        if ($this->checkLengthString($sRGB, 6))
        {
            $a = 0x00;//toujours 0 pour WriteExcel
            $r = hexdec(substr($sRGB, 0, 2));
            $g = hexdec(substr($sRGB, 2, 2));
            $b = hexdec(substr($sRGB, 4, 2));
            
            return $this->_convertRGBToColorIndex($r, $g, $b, $a);
        }
        else
            return FALSE;
    }
    
    /*
     * Converts the four color components R, G, B et A to index of Excel 97 palette
     * if it exists
     * 
     * @param int $r
     * @param int $g
     * @param int $b
     * @param int $a
     * @return int Color index
     */
    protected function _convertRGBToColorIndex($r, $g, $b, $a)
    {
        $bColorExistsInPalette = false;
        $hColorIndex = 0x00;
        
        foreach (static::$_XL97_COLOR_PALETTE as $hPaletteColorIndex => $aRGBA)
        {
            if ($aRGBA[0] == $r && $aRGBA[1] == $g && $aRGBA[2] == $b && $aRGBA[3] == $a)
            {
                $hColorIndex = $hPaletteColorIndex;
                $bColorExistsInPalette = true;
            }
        }
        
        if ($bColorExistsInPalette)
            return $hColorIndex;
        else
            return FALSE;
    }
    
    /*
     * Calculates the exact width of a column after a field content and the corresponding 
     * TrueType font. Modeled from PHPExcel_Shared_Font::calculateColumnWidth.
     * 
     * @param string $sFontName Font name of the style applied to the column
     * @param bool $bFontBold Bold of the font of the style applied to the column
     * @param bool $bFontItalic Italic of the font of the style applied to the column
     * @param float $fFontSize Size of the font of the style applied to the column
     * @param string $sCellText Cell content of the column
     * @param string $fRotation Rotation angle of the characters of the column field
     * @param string $sDefaultFontName Default font name if not set
     * @param float $fDefaultFontSize Size of the default font if not set
     * @return float Taille de la colonne
     */
    public function calculateExactColumnWidth($sFontName = '', $bFontBold = FALSE, $bFontItalic = FALSE, $fFontSize = 0.0, $sCellText = '', $fRotation = 0.0, $sDefaultFontName = '',  $fDefaultFontSize = 0.0)
    {
        //Special case if there are many carriage return characters
        
        if (FALSE !== strpos($sCellText, "\n")) {
            $aLineTexts = explode("\n", $sCellText);
            $aLineWidths = array();
            foreach ($aLineTexts as $sLineText) {
                $aLineWidths[] = $this->calculateExactColumnWidth($sFontName, $bFontBold, $bFontItalic, $fFontSize, $sLineText, $fRotation, $sDefaultFontName,  $fDefaultFontSize);
            }
            return max($aLineWidths); // width of the widest line
        }
        
        //Calculation of the exact width of the text in pixels
        
        //Width of the text including the padding
        $pColumnWidth = $this->getTextWidthPixelsExact($sCellText, $sFontName, $bFontBold, $bFontItalic, $fFontSize, $fRotation);
        
        //Excel adds some padding, use of width 1.07 of a 'n' glyph
        $pColumnWidth += ceil($this->getTextWidthPixelsExact('0', $sFontName, $bFontBold, $bFontItalic, $fFontSize, 0) * 1.07); // pixels incl. padding
        
        //Conversion of width in pixels to a column width
        $pColumnWidth = $this->pixelsToCellDimension($pColumnWidth, $sDefaultFontName,  $fDefaultFontSize);
        
        return round($pColumnWidth, 6);
    }
    
    /*
     * Calculates the exact width of a text in pixels after its TrueType font.
     * Modeled from PHPExcel_Shared_Font::getTextWidthPixelsExact.
     * 
     * @param string $sText Text
     * @param string $sFontName Font name of the text style
     * @param bool $bFontBold Bold of the font of the text style
     * @param bool $bFontItalic Italic of the font of the text style
     * @param float $fFontSize Size of the font of the text style
     * @param string $fRotation Rotation angle of the text characters
     * @return int Text size in pixels
     */
    public function getTextWidthPixelsExact($sText = '', $sFontName = '', $bFontBold = FALSE, $bFontItalic = FALSE, $fFontSize = 0.0, $fRotation = 0.0) {
        try {
            if (!function_exists('imagettfbbox')) {
                throw new \Exception(static::$__CLASS__ . ' -> getTextWidthPixelsExact - La librairie GD doit être activée.');
            }
            
            //The size of a font should be provided in pixels by GD2, but because
            //GD2 seems to be in 72dpi, the pixels and the points are equivalent
            $sFontFile = $this->getTrueTypeFontFileFromFont($sFontName, $bFontBold, $bFontItalic);
            $aTextBox = imagettfbbox($fFontSize, $fRotation, $sFontFile, $sText);
            
            //Obtaining the corner positions
            $iLowerLeftCornerX  = $aTextBox[0];
            $iLowerLeftCornerY  = $aTextBox[1];
            $iLowerRightCornerX = $aTextBox[2];
            $iLowerRightCornerY = $aTextBox[3];
            $iUpperRightCornerX = $aTextBox[4];
            $iUpperRightCornerY = $aTextBox[5];
            $iUpperLeftCornerX  = $aTextBox[6];
            $iUpperLeftCornerY  = $aTextBox[7];
            
            //Taking in account the rotation in the width calculation
            $iTextWidth = max($iLowerRightCornerX - $iUpperLeftCornerX, $iUpperRightCornerX - $iLowerLeftCornerX);

            return $iTextWidth;
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }

    /*
     * Obtaining the path of the TrueType file after a font. Modeled from 
     * PHPExcel_Shared_Font::getTrueTypeFontFileFromFont.
     * 
     * @param string $sFontName Font name of the text style
     * @param bool $bFontBold Bold of the font of the text style
     * @param bool $bFontItalic Italic of the font of the text style
     * @return string Path of the TTF file
     */
    public function getTrueTypeFontFileFromFont($sFontName = '', $bFontBold = FALSE, $bFontItalic = FALSE) {
        try {
            if (!file_exists(static::$_LIBRARY_PATH . self::PATH_TRUETYPE_FONTS) || !is_dir(static::$_LIBRARY_PATH . self::PATH_TRUETYPE_FONTS)) {
                throw new \Exception(static::$__CLASS__ . ' -> getTrueTypeFontFileFromFont - The directory of the TrueType font hasn\'t been properly defined.');
            }
            
            //Verification that we can find a font file
            switch ($sFontName) {
                case 'Arial':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_ARIAL_BOLD_ITALIC : self::FONT_FILENAME_ARIAL_BOLD) : ($bFontItalic ? self::FONT_FILENAME_ARIAL_ITALIC : self::FONT_FILENAME_ARIAL)
                            );
                    break;
                
                case 'Calibri':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_CALIBRI_BOLD_ITALIC : self::FONT_FILENAME_CALIBRI_BOLD) : ($bFontItalic ? self::FONT_FILENAME_CALIBRI_ITALIC : self::FONT_FILENAME_CALIBRI)
                            );
                    break;
                
                case 'Courier New':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_COURIER_NEW_BOLD_ITALIC : self::FONT_FILENAME_COURIER_NEW_BOLD) : ($bFontItalic ? self::FONT_FILENAME_COURIER_NEW_ITALIC : self::FONT_FILENAME_COURIER_NEW)
                            );
                    break;
                
                case 'Comic Sans MS':
                    $sFontFile = (
                            $bFontBold ? self::FONT_FILENAME_COMIC_SANS_MS_BOLD : self::FONT_FILENAME_COMIC_SANS_MS
                            );
                    break;
                
                case 'Georgia':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_GEORGIA_BOLD_ITALIC : self::FONT_FILENAME_GEORGIA_BOLD) : ($bFontItalic ? self::FONT_FILENAME_GEORGIA_ITALIC : self::FONT_FILENAME_GEORGIA)
                            );
                    break;
                
                case 'Impact':
                    $sFontFile = self::FONT_FILENAME_IMPACT;
                    break;
                
                case 'Liberation Sans':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_LIBERATION_SANS_BOLD_ITALIC : self::FONT_FILENAME_LIBERATION_SANS_BOLD) : ($bFontItalic ? self::FONT_FILENAME_LIBERATION_SANS_ITALIC : self::FONT_FILENAME_LIBERATION_SANS)
                            );
                    break;
                
                case 'Lucida Console':
                    $sFontFile = self::FONT_FILENAME_LUCIDA_CONSOLE;
                    break;
                
                case 'Lucida Sans Unicode':
                    $sFontFile = self::FONT_FILENAME_LUCIDA_SANS_UNICODE;
                    break;
                
                case 'Microsoft Sans Serif':
                    $sFontFile = self::FONT_FILENAME_MICROSOFT_SANS_SERIF;
                    break;
                
                case 'Palatino Linotype':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_PALATINO_LINOTYPE_BOLD_ITALIC : self::FONT_FILENAME_PALATINO_LINOTYPE_BOLD) : ($bFontItalic ? self::FONT_FILENAME_PALATINO_LINOTYPE_ITALIC : self::FONT_FILENAME_PALATINO_LINOTYPE)
                            );
                    break;
                
                case 'Symbol':
                    $sFontFile = self::FONT_FILENAME_SYMBOL;
                    break;
                
                case 'Tahoma':
                    $sFontFile = (
                            $bFontBold ? self::FONT_FILENAME_TAHOMA_BOLD : self::FONT_FILENAME_TAHOMA
                            );
                    break;
                
                case 'Times New Roman':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_TIMES_NEW_ROMAN_BOLD_ITALIC : self::FONT_FILENAME_TIMES_NEW_ROMAN_BOLD) : ($bFontItalic ? self::FONT_FILENAME_TIMES_NEW_ROMAN_ITALIC : self::FONT_FILENAME_TIMES_NEW_ROMAN)
                            );
                    break;
                
                case 'Trebuchet MS':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_TREBUCHET_MS_BOLD_ITALIC : self::FONT_FILENAME_TREBUCHET_MS_BOLD) : ($bFontItalic ? self::FONT_FILENAME_TREBUCHET_MS_ITALIC : self::FONT_FILENAME_TREBUCHET_MS)
                            );
                    break;
                
                case 'Verdana':
                    $sFontFile = (
                            $bFontBold ? ($bFontItalic ? self::FONT_FILENAME_VERDANA_BOLD_ITALIC : self::FONT_FILENAME_VERDANA_BOLD) : ($bFontItalic ? self::FONT_FILENAME_VERDANA_ITALIC : self::FONT_FILENAME_VERDANA)
                            );
                    break;
                
                default:
                    throw new \Exception(static::$__CLASS__ . ' -> getTrueTypeFontFileFromFont - Font name "' . $sFontName . '" unknown. Impossible to find a corresponding TrueType file.');
                    break;
            }

            $sFontFile = static::$_LIBRARY_PATH . self::PATH_TRUETYPE_FONTS . $sFontFile;

            //Verification of the existence of the font file
            if (!file_exists($sFontFile)) {
                throw new \Exception(static::$__CLASS__ . ' -> getTrueTypeFontFileFromFont - File of TrueType font not found.');
            }
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }

        return $sFontFile;
    }
    
    /*
     * Converts a dimension expressed in pixels to a column dimension.
     * Modeled from PHPExcel_Shared_Drawing::pixelsToCellDimension.
     * 
     * @param int $iValue Value in pixels
     * @param string $sFontName Font name
     * @param float $fFontSize Font size
     * @return float Column dimension
     */
    public function pixelsToCellDimension($iValue = 0, $sFontName = '', $fFontSize = 0.0) {
        if (isset(static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS[$sFontName][$fFontSize])) {
            //The exact width can be determined
            $fColWidth = $iValue * static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS[$sFontName][$fFontSize]['width'] / static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS[$sFontName][$fFontSize]['px'];
        } else {
            //We haven't any data for this precise font and size, then we use
            //an extrapolation of Calibri 11
            $fColWidth = $iValue * 11 * static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS['Calibri'][11]['width'] / static::$_XL2007_DEFAULT_FONT_COLUMN_WIDTHS['Calibri'][11]['px'] / $fFontSize;
        }
        
        return $fColWidth;
    }

    /*
     * Checks the type of a worksheet of an Excel workbook
     * 
     * It is considered that a workbook worksheet is necessarily an object, it can
     * however be reconsidered for other libraries.
     * 
     * @param object $oWorksheet Workbook workhsset to check
     * @return bool
     */
    public function checkWorksheetClass($oWorksheet)
    {
        try {
            if (is_object($oWorksheet) && get_class($oWorksheet) === static::$_WORKSHEET_CLASS)
            {
                return TRUE;
            }
            else
                throw new \Exception(static::$__CLASS__ . ' -> checkWorksheetClass - The worksheet type transmitted in parameter isn\'t correct.');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
    
    /*
     * Checks the type of the Excel style
     * 
     * It is considered that a style is necessarily an object, that can however
     * be reconsidered for other libraries.
     * 
     * @param object $oStyle Style à vérifier
     * @return bool
     */
    public function checkStyleClass($oStyle)
    {
        try {
            if (is_object($oStyle) && get_class($oStyle) === static::$_STYLE_CLASS)
            {
                return TRUE;
            }
            else
                throw new \Exception(static::$__CLASS__ . ' -> checkFormatClass - The style type transmitted in paramter isn\'t correct.');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
    
    /*
     * Checks the type and length of a string
     * 
     * @param string $sString String to check
     * @param int $iLength String length
     * @return bool
     */
    public function checkLengthString($sString, $iLength)
    {
        try {
            if (is_string($sString) && strlen($sString) === $iLength)
            {
                return TRUE;
            }
            else
                throw new \Exception(static::$__CLASS__ . ' -> checkLengthString - The transmitted parameter isn\'t a string of length ' . $iLength . '.');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
}