<?php

namespace MultiLibExcelExport\Excel\Adapter;

use MultiLibExcelExport\Excel\Adapter\Excel_Adapter;
use MultiLibExcelExport\Excel\Adapter\Excel_WorkbookAdapterInterface;

class Excel_SpreadsheetWriteExcel_WorkbookAdapter extends Excel_Adapter implements Excel_WorkbookAdapterInterface
{
    static protected $_INCLUDES = array();
    
    const PATH_LIBRARY = '/Spreadsheet-WriteExcel-2.40/';
    
    static protected $_WORKSHEET_CLASS;
    static protected $_STYLE_CLASS;
    
    protected  $_oWriteExcelWorkbook;
    
    protected $_oPerl;
    protected $_iNextStyle;
    
    protected $_fDefaultColSize = parent::DEFAULT_COL_SIZE;
    
    /* 
     * Constructor
     * 
     * @param string $sWorkbookFilename File path under which to save the Excel workbook
     * @param string $sLibraryPath Path to the Excel libraries
     */
    public function __construct($sWorkbookFilename, $sLibraryPath)
    {
        self::$__CLASS__ = __CLASS__;
        self::$_INCLUDES = array();
        
        self::$_WORKSHEET_CLASS = 'Perl::Spreadsheet::WriteExcel::Worksheet';
        self::$_STYLE_CLASS = 'Perl::Spreadsheet::WriteExcel::Format';
        
        $this->_iNextStyle = 0;
        $this->_fDefaultColSize = parent::DEFAULT_COL_SIZE;
        
        parent::__construct($sLibraryPath);
        
        try {
            $this->_oPerl = new \Perl();
            
            $sLibraryPath = str_replace('\\', '/', static::$_LIBRARY_PATH . self::PATH_LIBRARY);
            
            $this->_oPerl->eval(
<<<PERL
use strict;
BEGIN{ @INC = ( "$sLibraryPath" , @INC ); }
use Spreadsheet::WriteExcel;
use Encode 'decode';
PERL
            );
            
            $sWorkbookFilename = str_replace('\\', '/', $sWorkbookFilename);
            
            $this->_oPerl->eval(
<<<PERL
\$workbook = Spreadsheet::WriteExcel->new(qq($sWorkbookFilename));
PERL
            );
        
        } catch (\PerlException $e) {
            echo "Perl error __construct: " . $e->getMessage() . "\n";
            return FALSE;
        }
    }
    
    /*
     * Adds a worksheeet to the Excel workbook
     * 
     * @param string $sWorksheetName Name of the worksheeet to add
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     * @return int Index of the worksheet in the workbook
     */
    public function addWorksheetToWorkbook($sWorksheetName = '', $iIndexInWorkBook = 0)
    {
        try {
            $this->_oPerl->tmpWorksheetName = $sWorksheetName;
            $this->_oPerl->eval(
<<<PERL
\$tmpWorksheetName = decode('utf-8', \$tmpWorksheetName);
\$worksheet{$iIndexInWorkBook} = \$workbook->add_worksheet(qq/\$tmpWorksheetName/);
\$worksheet{$iIndexInWorkBook}->keep_leading_zeros();
PERL
            );
            unset($this->_oPerl->tmpWorksheetName);
        } catch (\PerlException $e) {
            echo "Perl error addWorksheetToWorkbook: " . $e->getMessage() . "\n";
            return FALSE;
        }
        return $iIndexInWorkBook;
    }
    
    /*
     * Sets the workhsset as active in the Excel workbook
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     */
    public function setWorksheetAsActiveInWorkbook($iWorksheet, $iIndexInWorkBook)
    {
        try {
            $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->activate();
PERL
            );
        } catch (\PerlException $e) {
            echo "Perl error setWorksheetAsActiveInWorkbook: " . $e->getMessage() . "\n";
            return FALSE;
        }
    }
    
    /*
     * Converts a style into a Spreadsheet::WriteExcel style
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
     * numberformat => code (format chaîne ou index)
     * alignment    => horizontal
     *              => vertical
     *              => wrap
     * borders      => allborders => color => argb
     *                                     => rgb
     *                                     => index
     *                            => style
     * 
     * @param int $iStyle Index of style
     * @param array $aStyle Style under array form
     * @return array|bool Style converted to an array that can be added to workbook
     * Spreadsheet::WriteExcel via the add_format() method or false
     */
    protected function _convertAStyleToOStyleForWriteExcel($iStyle, $aStyle = array())
    {
        if (!empty($aStyle)) 
        {
            try {
                foreach ($aStyle as $sKey => $value) {
                    switch ($sKey) {
                        case 'font':
                            if (!empty($value))
                                foreach ($value as $sKey2 => $value2) {
                                    switch ($sKey2) {
                                        case 'name':
                                            switch ($value2) {//comme dans PHPExcel_Shared_Font->getCharsetFromFontName()
                                                case 'EucrosiaUPC': $this->_oPerl->eval("\$style{$iStyle}->set_font_charset(0xDD);");
                                                    break;
                                                case 'Wingdings': $this->_oPerl->eval("\$style{$iStyle}->set_font_charset(0x02);");
                                                    break;
                                                case 'Wingdings 2': $this->_oPerl->eval("\$style{$iStyle}->set_font_charset(0x02);");
                                                    break;
                                                case 'Wingdings 3': $this->_oPerl->eval("\$style{$iStyle}->set_font_charset(0x02);");
                                                    break;
                                                default: $this->_oPerl->eval("\$style{$iStyle}->set_font_charset(0x00);");
                                                    break;
                                            }
                                            $this->_oPerl->eval("\$style{$iStyle}->set_font('{$value2}');");
                                            break;
                                        case 'size':
                                            $this->_oPerl->eval("\$style{$iStyle}->set_size({$value2});");
                                            break;
                                        case 'bold':
                                            switch ($value2) {
                                                case TRUE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_bold(1);");
                                                    break;
                                                case FALSE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_bold(0);");
                                                    break;
                                            }
                                            break;
                                        case 'italic':
                                            switch ($value2) {
                                                case TRUE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_italic(1);");
                                                    break;
                                                case FALSE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_italic(0);");
                                                    break;
                                            }
                                            break;
                                        case 'color':
                                            if (!empty($value2))
                                                foreach ($value2 as $sKey3 => $value3) {
                                                    switch ($sKey3) {
                                                        case 'argb':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_color({$this->convertARGBToColorIndex($value3)});");
                                                            break;
                                                        case 'rgb':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_color({$this->convertRGBToColorIndex($value3)});");
                                                            break;
                                                        case 'index':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_color({$value3});");
                                                            break;
                                                    }
                                                }
                                            break;
                                        case 'underline':
                                            switch ($value2) {
                                                case 'none':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_underline(0x00);");
                                                    break;
                                                case 'double':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_underline(0x02);");
                                                    break;
                                                case 'doubleAccounting':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_underline(0x22);");
                                                    break;
                                                case 'single':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_underline(0x01);");
                                                    break;
                                                case 'singleAccounting':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_underline(0x21);");
                                                    break;
                                            }
                                            break;
                                    }
                                }
                            break;
                        case 'fill':
                            if (!empty($value))
                                foreach ($value as $sKey2 => $value2) {
                                    switch ($sKey2) {
                                        case 'type':
                                            switch ($value2) {
                                                case 'none':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x00);");
                                                    break;
                                                case 'solid':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x01);");
                                                    break;
                                                case 'linear':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x00);");
                                                    break;
                                                case 'path':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x00);");
                                                    break;
                                                case 'darkDown':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x07);");
                                                    break;
                                                case 'darkGray':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x03);");
                                                    break;
                                                case 'darkGrid':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x09);");
                                                    break;
                                                case 'darkHorizontal':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x05);");
                                                    break;
                                                case 'darkTrellis':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0A);");
                                                    break;
                                                case 'darkUp':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x08);");
                                                    break;
                                                case 'darkVertical':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x06);");
                                                    break;
                                                case 'gray0625':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x12);");
                                                    break;
                                                case 'gray125':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x11);");
                                                    break;
                                                case 'lightDown':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0D);");
                                                    break;
                                                case 'lightGray':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x04);");
                                                    break;
                                                case 'lightGrid':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0F);");
                                                    break;
                                                case 'lightHorizontal':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0B);");
                                                    break;
                                                case 'lightTrellis':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x10);");
                                                    break;
                                                case 'lightUp':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0E);");
                                                    break;
                                                case 'lightVertical':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x0C);");
                                                    break;
                                                case 'mediumGray':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_pattern(0x02);");
                                                    break;
                                            }
                                            break;
                                        case 'startcolor':
                                            if (!empty($value2))
                                                foreach ($value2 as $sKey3 => $value3) {
                                                    switch ($sKey3) {
                                                        case 'argb':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_fg_color({$this->convertARGBToColorIndex($value3)});");
                                                            break;
                                                        case 'rgb':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_fg_color({$this->convertRGBToColorIndex($value3)});");
                                                            break;
                                                        case 'index':
                                                            $this->_oPerl->eval("\$style{$iStyle}->set_fg_color({$value3});");
                                                            break;
                                                    }
                                                }
                                            break;
                                    }
                                }
                            break;
                        case 'numberformat':
                            if (!empty($value))
                                foreach ($value as $sKey2 => $value2) {
                                    switch ($sKey2) {
                                        case 'code':
                                            $pNumFormat = $this->_convertNumFormatCodeToIndex($value2);
                                            $this->_oPerl->tmpNumFormat = $pNumFormat;
                                            $this->_oPerl->eval(
<<<PERL
\$style{$iStyle}->set_num_format(\$tmpNumFormat);
PERL
                                            );
                                            unset($this->_oPerl->tmpNumFormat);
                                            break;
                                    }
                                }
                            break;

                        case 'alignment':
                            if (!empty($value))
                                foreach ($value as $sKey2 => $value2) {
                                    switch ($sKey2) {
                                        case 'horizontal':
                                            switch ($value2) {
                                                case 'general':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(0);");
                                                    break;
                                                case 'left':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(1);");
                                                    break;
                                                case 'right':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(3);");
                                                    break;
                                                case 'center':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(2);");
                                                    break;
                                                case 'centerContinuous':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(6);");
                                                    break;
                                                case 'justify':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(5);");
                                                    break;
                                                case 'distributed':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(7);");
                                                    break;
                                                case 'fill':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_h_align(4);");
                                                    break;
                                            }
                                            break;
                                        case 'vertical':
                                            switch ($value2) {
                                                case 'bottom':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_v_align(2);");
                                                    break;
                                                case 'top':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_v_align(0);");
                                                    break;
                                                case 'center':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_v_align(1);");
                                                    break;
                                                case 'justify':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_v_align(3);");
                                                    break;
                                                case 'distributed':
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_v_align(4);");
                                                    break;
                                            }
                                            break;
                                        case 'wrap':
                                            switch ($value2) {
                                                case TRUE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_wrap(1);");
                                                    break;
                                                case FALSE:
                                                    $this->_oPerl->eval("\$style{$iStyle}->set_text_wrap(0);");
                                                    break;
                                            }
                                            break;
                                    }
                                }
                            break;

                        case 'borders':
                            if (!empty($value))
                                foreach ($value as $sKey2 => $value2) {
                                    switch ($sKey2) {
                                        case 'allborders':
                                            if (!empty($value2))
                                                foreach ($value2 as $sKey3 => $value3) {
                                                    switch ($sKey3) {
                                                        case 'color':
                                                            if (!empty($value3))
                                                                foreach ($value3 as $sKey4 => $value4) {
                                                                    switch ($sKey4) {
                                                                        case 'argb':
                                                                            $this->_oPerl->eval("\$style{$iStyle}->set_border_color({$this->convertARGBToColorIndex($value4)});");
                                                                            break;
                                                                        case 'rgb':
                                                                            $this->_oPerl->eval("\$style{$iStyle}->set_border_color({$this->convertRGBToColorIndex($value4)});");
                                                                            break;
                                                                        case 'index':
                                                                            $this->_oPerl->eval("\$style{$iStyle}->set_border_color({$value4});");
                                                                            break;
                                                                    }
                                                                }
                                                            break;
                                                        case 'style':
                                                            switch ($value3) {
                                                                case 'none':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x00);");
                                                                    break;
                                                                case 'dashDot':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x09);");
                                                                    break;
                                                                case 'dashDotDot':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x0B);");
                                                                    break;
                                                                case 'dashed':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x03);");
                                                                    break;
                                                                case 'dotted':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x04);");
                                                                    break;
                                                                case 'double':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x06);");
                                                                    break;
                                                                case 'hair':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x07);");
                                                                    break;
                                                                case 'medium':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x02);");
                                                                    break;
                                                                case 'mediumDashDot':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x0A);");
                                                                    break;
                                                                case 'mediumDashDotDot':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x0C);");
                                                                    break;
                                                                case 'mediumDashed':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x08);");
                                                                    break;
                                                                case 'slantDashDot':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x0D);");
                                                                    break;
                                                                case 'thick':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x05);");
                                                                    break;
                                                                case 'thin':
                                                                    $this->_oPerl->eval("\$style{$iStyle}->set_border(0x01);");
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
            } catch (\PerlException $e) {
                echo "Perl error _convertAStyleToOStyleForWriteExcel: " . $e->getMessage() . "\n";
                return FALSE;
            }
        }
        
        return $iStyle;
    }
    
    /*
     * Adds a style to the Excel workbook
     * 
     * @param $aStyle Style under array form
     * @return int|bool Index of added style or false
     */
    public function addStyleToWorkbook(&$aStyle = array())
    {
        try {
            $this->_oPerl->eval(
<<<PERL
\$style{$this->_iNextStyle} = \$workbook->add_format();
PERL
            );
            $this->_iNextStyle++;
        } catch (\PerlException $e) {
            echo "Perl error addStyleToWorkbook: " . $e->getMessage() . "\n";
            return FALSE;
        }
        
        $this->_convertAStyleToOStyleForWriteExcel($this->_iNextStyle-1, $aStyle);
        
        return $this->_iNextStyle-1;
    }
    
    /*
     * Closes the Excel workbook saving it
     */
    public function closeWorkbook()
    {
        try {
            $this->_oPerl->eval(
<<<PERL
\$workbook->close();
PERL
            );
        } catch (\PerlException $e) {
            echo "Perl error closeWorkbook: " . $e->getMessage() . "\n";
            return FALSE;
        }
    }
    
    /*
     * Sets the width of one or many columns of an Excel worksheet
     * 
     * @param int $iWorksheet Index de la feuille dans le classeur
     * @param float $fWidth Width
     * @param int $iIndexInWorkbook Index of the workbook worksheet in façade
     * @param int $iFirstcol Index of the first column
     * @param int $iLastcol Index of the last column
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function setColumnsWidthToWorksheet($iWorksheet, $fWidth, $iIndexInWorkbook = 0, $iFirstcol = 0, $iLastcol = 0)
    {
        if (is_int($iWorksheet))
        {
            try {
                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->set_column({$iFirstcol}, {$iLastcol}, {$fWidth});
PERL
                );

            } catch (\PerlException $e) {
                echo "Perl error setColumnsWidthToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Insertion of an image into an Excel worksheet. Supports the types .bmp, 
     * .png, .jpg
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param string $sFilename File path of the image
     * @param int $iRow Index of the cell line where is inserted the image
     * @param int $iCol Index of the cell column where is inserted the image
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function insertImageToWorksheet($iWorksheet, $sFilename, $iRow = 0, $iCol = 0)
    {
        if (is_int($iWorksheet))
        {
            $sFilename = str_replace('\\', '/', $sFilename);
            
            try {
                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->insert_image({$iRow}, {$iCol}, '{$sFilename}');
PERL
                );
            } catch (\PerlException $e) {
                echo "Perl error insertBitmapToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Permits to calculate and specify the width of a column after its content
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iCol Index of the column
     * @param string $pCellToken Cell content of type NULL, string, int, float, double, bool
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of
     * the columns after the TrueType fonts
     */
    protected function _setColumnAutosizeToWorksheet($iWorksheet, $iCol = 0, $pCellToken = NULL, $iStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
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
        
        try {
            $this->_oPerl->eval(
<<<PERL
if  (!exists(\$worksheet{$iWorksheet}->{_col_sizes}->{{$iCol}}) or {$fCalculatedWidth} > \$worksheet{$iWorksheet}->{_col_sizes}->{{$iCol}})
{
    \$worksheet{$iWorksheet}->set_column({$iCol}, {$iCol}, {$fCalculatedWidth});
}
PERL
            );
        } catch (\PerlException $e) {
            echo "Perl error _setColumnAutosizeToWorksheet: " . $e->getMessage() . "\n";
            return FALSE;
        }
    }
    
    /*
     * Permits to calculate and specify the width of the columns of cells merged 
     * horizontally depending on the content of the upper left cell
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iFirstCol Index of the first column of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param $pCellToken Value of upper left cell of type NULL, string, int, float, double, bool
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of
     * the columns after the TrueType fonts
     */
    protected function _setMergedColumnsAutosizeToWorksheet($iWorksheet, $iFirstCol = 0, $iLastCol = 0, $pCellToken = NULL, $iStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
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
        
        try {
            $this->_oPerl->eval(
<<<PERL
\$iNbCol = {$iLastCol} - {$iFirstCol} + 1;
\$fSumColWidth = 0.0;
for (\$iCol = {$iFirstCol}; \$iCol <= {$iLastCol}; \$iCol++) {
    \$fSumColWidth += \$worksheet{$iWorksheet}->{_col_sizes}->{\$iCol};
}
if ({$fCalculatedWidth} > \$fSumColWidth)
{
    \$fRestWidthPerCol = ({$fCalculatedWidth} - \$fSumColWidth) / \$iNbCol;
    for (\$iCol = {$iFirstCol}; \$iCol <= {$iLastCol}; \$iCol++) {
        \$worksheet{$iWorksheet}->set_column(\$iCol, \$iCol, \$worksheet{$iWorksheet}->{_col_sizes}->{\$iCol} + \$fRestWidthPerCol);
    }
}
PERL
            );
        } catch (\PerlException $e) {
            echo "Perl error _setColumnAutosizeToWorksheet: " . $e->getMessage() . "\n";
            return FALSE;
        }
    }
    
    /*
     * Write into a cell range with an unique style. The style is always overwritten 
     * either if $iStyle is NULL or not.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line to write
     * @param int $iFirstCol Index of the first column to write
     * @param int $iLastRow Index of the last line of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param array $aCellMatrix Cell matrix of the form $aCellMatrix[$iRow][$iCol] = $pValeur
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param array $aMaxStringWithLengthPerCol Array indexed by column giving the maximal
     * length of the field of the column
     * @param array $aIndexColCellsWithString Index of the cells containing a
     * string
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeFromArrayToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), $iStyle = NULL, &$aStyle = array(), &$aMaxStringWithLengthPerCol = array(), &$aIndexColCellsWithString = array())
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                foreach ($aCellMatrix as $iTmpY => $aCol) {
                    if (!empty($aCol))
                    {
                        $this->_oPerl->array->aCol = $aCol;
                        $this->_oPerl->array->aIndexColCellsWithString = $aIndexColCellsWithString[$iTmpY];
                        
                        $this->_oPerl->eval(
<<<PERL
for my \$i (0 .. $#aIndexColCellsWithString) {
     \$aCol[\$aIndexColCellsWithString[\$i]] = decode('utf-8', \$aCol[\$aIndexColCellsWithString[\$i]]);                                     
}

\$worksheet{$iWorksheet}->write_col({$iFirstRow}, {$iFirstCol}, \@aCol, \$style{$iStyle});
PERL
                        );
                        unset($this->_oPerl->array->aCol);
                        unset($this->_oPerl->array->aIndexColCellsWithString);
                    }
                    ++$iFirstCol;
                }
            } catch (\PerlException $e) {
                echo "Perl error writeFromArrayToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            if (!empty($aMaxStringWithLengthPerCol)) foreach ($aMaxStringWithLengthPerCol as $iCol => $aStringLength)
            {
                if (!empty($aStringLength) && !empty($aCellMatrix[$iCol])) $this->_setColumnAutosizeToWorksheet($iWorksheet, $iCol, $aStringLength[2], $iStyle, $aStyle, $aStringLength[1]);
            }
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes with a style into a cell of an Excel worksheet. The style is always overwritten 
     * either if $iStyle is NULL or not.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pToken Content to write of type NULL, string, int, float, double, bool
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pToken = NULL, $iStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                $this->_oPerl->tmpToken = $pToken;
                    
                if (is_string($pToken))
                {
                    $this->_oPerl->eval(
<<<PERL
\$tmpToken = decode('utf-8', \$tmpToken);
PERL
                    );
                }

                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->write({$iRow}, {$iCol}, \$tmpToken, \$style{$iStyle});
PERL
                );
                
                unset($this->_oPerl->tmpToken);
            } catch (\PerlException $e) {
                echo "Perl error writeToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($iWorksheet, $iCol, $pToken, $iStyle, $aStyle);
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a string with a style into a cell of an Excel worksheet. The style 
     * is always overwritten either if $iStyle is NULL or not.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param string $sValue Content to write
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeStringToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $sValue = '', $iStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                $this->_oPerl->tmpValue = $sValue;
                $this->_oPerl->eval(
<<<PERL
\$tmpValue = decode('utf-8', \$tmpValue);
\$worksheet{$iWorksheet}->write_string({$iRow}, {$iCol}, \$tmpValue, \$style{$iStyle});
PERL
                );
                unset($this->_oPerl->tmpValue);
            } catch (\PerlException $e) {
                echo "Perl error writeStringToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($iWorksheet, $iCol, $sValue, $iStyle, $aStyle);
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a number with a style into a cell of an Excel worksheet. The style 
     * is always overwritten either if $iStyle is NULL or not.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pNumber Content to write of type string, int, float, double
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeNumberToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pNumber = 0, $iStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bMergedCell = FALSE)
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->write_number({$iRow}, {$iCol}, {$pNumber}, \$style{$iStyle});
PERL
                );
            } catch (\PerlException $e) {
                echo "Perl error writeNumberToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($iWorksheet, $iCol, $pNumber, $iStyle, $aStyle, $iNbCharPercent);
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes an empty cell with a style into an Excel worksheet. The style 
     * is always overwritten either if $iStyle is NULL or not.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeBlankToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $iStyle = NULL, &$aStyle = array())
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->write_blank({$iRow}, {$iCol}, \$style{$iStyle});
if (!exists(\$worksheet{$iWorksheet}->{_col_sizes}->{{$iCol}}))
{
    \$worksheet{$iWorksheet}->set_column({$iCol}, {$iCol}, {$this->_fDefaultColSize});
}
PERL
                );
            } catch (\PerlException $e) {
                echo "Perl error writeBlankToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes empty cells into an Excel worksheet. Style is overwritten only if
     * cells are already empty and without style.
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function writeBlankToManyCellsToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $iStyle = NULL, &$aStyle = array())
    {
        if (is_int($iWorksheet))
        {
            if (isset($iStyle) && !is_int($iStyle)) return FALSE;
            
            try {
                for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                    for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                        $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->write_blank({$iRow}, {$iCol}, \$style{$iStyle});
if (!exists(\$worksheet{$iWorksheet}->{_col_sizes}->{{$iCol}}))
{
    \$worksheet{$iWorksheet}->set_column({$iCol}, {$iCol}, {$this->_fDefaultColSize});
}
PERL
                        );
                    }
                }
            } catch (\PerlException $e) {
                echo "Perl error writeBlankToManyCellsToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Merge cells in an Excel worksheet
     * 
     * @param int $iWorksheet Index of the worksheet in the workbook
     * @param int $iIndexInWorkbook Index de la feuille du classeur en façade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * 
     * Only to resize the columns:
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param int $iStyle Index of style
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * 
     * @return int|bool Index of the worksheet in the workbook or false
     */
    public function mergeCellsToWorksheet($iWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, $iStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1)
    {
        if (is_int($iWorksheet))
        {
            try {
                $this->_oPerl->eval(
<<<PERL
\$worksheet{$iWorksheet}->merge_cells({$iFirstRow}, {$iFirstCol}, {$iLastRow}, {$iLastCol});
PERL
                );
            } catch (\PerlException $e) {
                echo "Perl error mergeCellsToWorksheet: " . $e->getMessage() . "\n";
                return FALSE;
            }
            if (!is_null($pCellToken) && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setMergedColumnsAutosizeToWorksheet($iWorksheet, $iFirstCol, $iLastCol, $pCellToken, $iStyle, $aStyle, $iNbCharPercent);
            
            return $iWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Converts a format code of an Excel cell into its integer index if it's part
     * of those by default in Excel
     * 
     * @param $pNumFormat Format code
     * @return Format code
     */
    protected function _convertNumFormatCodeToIndex($pNumFormat = '')
    {
        if (is_string($pNumFormat))
        {
            switch (true) {
                case $pNumFormat === 'General':
                    $pReturnedNumberFormat = 0x00;
                    break;
                case $pNumFormat === '0':
                    $pReturnedNumberFormat = 0x01;
                    break;
                case $pNumFormat === '0.00':
                    $pReturnedNumberFormat = 0x02;
                    break;
                case $pNumFormat === '#,##0':
                    $pReturnedNumberFormat = 0x03;
                    break;
                case $pNumFormat === '#,##0.00':
                    $pReturnedNumberFormat = 0x04;
                    break;
                case $pNumFormat === '0$;(0$)':
                    $pReturnedNumberFormat = 0x05;
                    break;
                case $pNumFormat === '0$;[Red](0$)':
                    $pReturnedNumberFormat = 0x06;
                    break;
                case $pNumFormat ===  '$0.00;($0.00)':
                    $pReturnedNumberFormat = 0x07;
                    break;
                case $pNumFormat === '$0.00;[Red]($0.00)':
                    $pReturnedNumberFormat = 0x08;
                    break;
                case $pNumFormat === '0%':
                    $pReturnedNumberFormat = 0x09;
                    break;
                case $pNumFormat === '0.00%':
                    $pReturnedNumberFormat = 0x0a;
                    break;
                case $pNumFormat === '0.00E+00':
                    $pReturnedNumberFormat = 0x0b;
                    break;
                case $pNumFormat === '#" "?/?':
                    $pReturnedNumberFormat = 0x0c;
                    break;
                case $pNumFormat === '#" "??/??':
                    $pReturnedNumberFormat = 0x0d;
                    break;
                case $pNumFormat === 'mm-dd-yy':
                    $pReturnedNumberFormat = 0x0e;
                    break;
                case $pNumFormat === 'd-mmm-yy':
                    $pReturnedNumberFormat = 0x0f;
                    break;
                case $pNumFormat === 'd-mmm':
                    $pReturnedNumberFormat = 0x10;
                    break;
                case $pNumFormat === 'mmm-yy':
                    $pReturnedNumberFormat = 0x11;
                    break;
                case $pNumFormat === 'h:mm AM/PM':
                    $pReturnedNumberFormat = 0x12;
                    break;
                case $pNumFormat === 'h:mm:ss AM/PM':
                    $pReturnedNumberFormat = 0x13;
                    break;
                case $pNumFormat === 'h:mm':
                    $pReturnedNumberFormat = 0x14;
                    break;
                case $pNumFormat === 'h:mm:ss':
                    $pReturnedNumberFormat = 0x15;
                    break;
                case $pNumFormat === 'm/d/yy h:mm':
                    $pReturnedNumberFormat = 0x16;
                    break;
                case $pNumFormat === '#,##0;(#,##0)':
                    $pReturnedNumberFormat = 0x25;
                    break;
                case $pNumFormat === '#,##0;[Red](#,##0)':
                    $pReturnedNumberFormat = 0x26;
                    break;
                case $pNumFormat === '#,##0.00;(#,##0.00)':
                    $pReturnedNumberFormat = 0x27;
                    break;
                case $pNumFormat === '#,##0.00;[Red](#,##0.00)':
                    $pReturnedNumberFormat = 0x28;
                    break;
                case$pNumFormat ===  '_-#,##0_-;_-"-"#,##0_-;_-@_-':
                    $pReturnedNumberFormat = 0x29;
                    break;
                case $pNumFormat === '_-$* #,##0_-;_-$* "-"#,##0_-;_-$* "-"_-;_-@_-':
                    $pReturnedNumberFormat = 0x2a;
                    break;
                case $pNumFormat === '_-#,##0.00_-;_-"-"#,##0.00_-;_-"-"??_-;_-@_-':
                    $pReturnedNumberFormat = 0x2b;
                    break;
                case $pNumFormat === '_-$* #,##0.00_-;_-$* "-"#,##0.00_-;_-$* "-"??_-;_-@_-':
                    $pReturnedNumberFormat = 0x2c;
                    break;
                case $pNumFormat === 'mm:ss':
                    $pReturnedNumberFormat = 0x2d;
                    break;
                case $pNumFormat === '[h]:mm:ss':
                    $pReturnedNumberFormat = 0x2e;
                    break;
                case $pNumFormat === 'mmss.0':
                    $pReturnedNumberFormat = 0x2f;
                    break;
                case $pNumFormat === '##0.0E+0':
                    $pReturnedNumberFormat = 0x30;
                    break;
                case $pNumFormat === '@':
                    $pReturnedNumberFormat = 0x31;
                    break;
                default:
                    $pReturnedNumberFormat = $pNumFormat;
                    break;
            }
        }
        else
            $pReturnedNumberFormat = (int)$pNumFormat;
            
        return $pReturnedNumberFormat;
    }
    
    /*
     * Sets a number format in an Excel style
     * 
     * @param int $iStyle Index of style
     * @param $pNumFormat Number format of type string or integer for index
     * @return int|bool Index of style or false
     */
    public function setNumFormatToStyle($iStyle, $pNumFormat = '')
    {
        if (is_int($iStyle))
        {
            try {
                $pNumFormat = $this->_convertNumFormatCodeToIndex($pNumFormat);
                $this->_oPerl->tmpNumFormat = $pNumFormat;
                $this->_oPerl->eval(
<<<PERL
\$style{$iStyle}->set_num_format(\$tmpNumFormat);
PERL
                );
                unset($this->_oPerl->tmpNumFormat);
            } catch (\PerlException $e) {
                echo "Perl error setNumFormatToStyle: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iStyle;
        }
        else
            return FALSE;
    }
    
    /*
     * Specifies the text wrap in an Excel style
     * 
     * @param int $iStyle Index of style
     * @return int|bool Index of style or false
     */
    public function setTextWrapToStyle($iStyle)
    {
        if (is_int($iStyle))
        {
            try {
                $this->_oPerl->eval(
<<<PERL
\$style{$iStyle}->set_text_wrap();
PERL
                );
            } catch (\PerlException $e) {
                echo "Perl error setTextWrapToStyle: " . $e->getMessage() . "\n";
                return FALSE;
            }
            
            return $iStyle;
        }
        else
            return FALSE;
    }
}