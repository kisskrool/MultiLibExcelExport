<?php

namespace MultiLibExcelExport\Excel\Adapter;

use MultiLibExcelExport\Excel\Adapter\Excel_Adapter;
use MultiLibExcelExport\Excel\Adapter\Excel_WorkbookAdapterInterface;

class Excel_WriteExcel_WorkbookAdapter extends Excel_Adapter implements Excel_WorkbookAdapterInterface
{
    static protected $_INCLUDES;
    
    static protected $_WORKSHEET_CLASS;
    static protected $_STYLE_CLASS;
    
    protected $_writeexcel_workbook;
    protected $_sWorkbookFilename;
    
    protected $_fDefaultColSize;
    
    /* 
     * Constructor
     * 
     * @param string $sWorkbookFilename File path under which to save the Excel workbook
     * @param string $sLibraryPath Path to the Excel libraries
     */
    public function __construct($sWorkbookFilename, $sLibraryPath)
    {
        self::$__CLASS__ = __CLASS__;        
        self::$_INCLUDES = array('/php_writeexcel-0.3.0/class.writeexcel_workbook.inc.php',
                                 '/php_writeexcel-0.3.0/class.writeexcel_workbookbig.inc.php',
                                 '/php_writeexcel-0.3.0/class.writeexcel_worksheet.inc.php',
                                 );
        
        self::$_WORKSHEET_CLASS = 'writeexcel_worksheet';
        self::$_STYLE_CLASS = 'writeexcel_format';
        
        $this->_fDefaultColSize = parent::DEFAULT_COL_SIZE;
        
        parent::__construct($sLibraryPath);
        
        $this->_writeexcel_workbook = new \writeexcel_workbookbig($sWorkbookFilename);
        $this->_sWorkbookFilename = $sWorkbookFilename;
    }
    
    /*
     * Sanitizes the string containing the Euro symbol
     * 
     * @param string $sChaine String to sanitize
     * @return string Processed string
     */
    protected function _sanitizeUTF8($sString)
    {
        $sString = iconv("utf-8", "utf-8//IGNORE", $sString);
        $sString = iconv("UTF-8", "CP1252", $sString);
        return $sString;
    }
    
    /*
     * Adds a worksheeet to the Excel workbook
     * 
     * @param string $sWorksheetName Name of the worksheeet to add
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     * @return writeexcel_worksheet Worksheet of type WriteExcel
     */
    public function addWorksheetToWorkbook($sWorksheetName = '', $iIndexInWorkBook = 0)
    {
        $oWorksheet = $this->_writeexcel_workbook->addworksheet($this->_sanitizeUTF8($sWorksheetName));
        
        return $oWorksheet;
    }
    
    /*
     * Sets the workhsset as active in the Excel workbook
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     */
    public function setWorksheetAsActiveInWorkbook($oWorksheet, $iIndexInWorkBook)
    {
        $this->_writeexcel_workbook->_activesheet = $oWorksheet->_index;
        $oWorksheet->_selected = 1;
    }
    
    /*
     * Converts a style into a WriteExcel style
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
     * @param array $aStyle Style under array form
     * @return array Style converted to an array that can be added to workbook
     * WriteExcel via the addformat() method
     */
    protected function _convertAStyleForWriteExcel($aStyle = array())
    {
        $aConvertedAStyle = array();
        
        if (!empty($aStyle)) foreach ($aStyle as $sKey => $value)
        {
            switch($sKey) {
                case 'font':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {
                            case 'name':
                                switch ($value2) {//like in PHPExcel_Shared_Font->getCharsetFromFontName()
                                    case 'EucrosiaUPC': $aConvertedAStyle['font_charset'] = 0xDD; break;
                                    case 'Wingdings':	$aConvertedAStyle['font_charset'] = 0x02; break;
                                    case 'Wingdings 2':	$aConvertedAStyle['font_charset'] = 0x02; break;
                                    case 'Wingdings 3':	$aConvertedAStyle['font_charset'] = 0x02; break;
                                    default:            $aConvertedAStyle['font_charset'] = 0x00; break;
                                }
                                $aConvertedAStyle['font'] = $value2;
                                break;
                            case 'size':
                                $aConvertedAStyle['size'] = $value2;
                                break;
                            case 'bold':
                                switch($value2) {
                                    case TRUE:
                                        $aConvertedAStyle['bold'] = TRUE;
                                        break;
                                    case FALSE:
                                        $aConvertedAStyle['bold'] = FALSE;
                                        break;
                                }
                                break;
                            case 'italic':
                                switch($value2) {
                                    case TRUE:
                                        $aConvertedAStyle['italic'] = TRUE;
                                        break;
                                    case FALSE:
                                        $aConvertedAStyle['italic'] = FALSE;
                                        break;
                                }
                                break;
                            case 'color':
                                if (!empty($value2)) foreach ($value2 as $sKey3 => $value3)
                                {
                                    switch($sKey3) {
                                        case 'argb':
                                            $aConvertedAStyle['color'] = $this->convertARGBToColorIndex($value3);
                                            break;
                                        case 'rgb':
                                            $aConvertedAStyle['color'] = $this->convertRGBToColorIndex($value3);
                                            break;
                                        case 'index':
                                            $aConvertedAStyle['color'] = $value3;
                                            break;
                                    }
                                }
                                break;
                            case 'underline':
                                switch($value2) {
                                    case 'none':
                                        $aConvertedAStyle['underline'] = 0x00;
                                        break;
                                    case 'double':
                                        $aConvertedAStyle['underline'] = 0x02;
                                        break;
                                    case 'doubleAccounting':
                                        $aConvertedAStyle['underline'] = 0x22;
                                        break;
                                    case 'single':
                                        $aConvertedAStyle['underline'] = 0x01;
                                        break;
                                    case 'singleAccounting':
                                        $aConvertedAStyle['underline'] = 0x21;
                                        break;
                                }
                                break;
                        }
                    }
                    break;
                case 'fill':
                    if (!empty($value)) foreach ($value as $sKey2 => $value2)
                    {
                        switch($sKey2) {                            
                            case 'type':
                                switch($value2) {
                                    case 'none':
                                        $aConvertedAStyle['pattern'] = 0x00;
                                        break;
                                    case 'solid':
                                        $aConvertedAStyle['pattern'] = 0x01;
                                        break;
                                    case 'linear':
                                        $aConvertedAStyle['pattern'] = 0x00;
                                        break;
                                    case 'path':
                                        $aConvertedAStyle['pattern'] = 0x00;
                                        break;
                                    case 'darkDown':
                                        $aConvertedAStyle['pattern'] = 0x07;
                                        break;
                                    case 'darkGray':
                                        $aConvertedAStyle['pattern'] = 0x03;
                                        break;
                                    case 'darkGrid':
                                        $aConvertedAStyle['pattern'] = 0x09;
                                        break;
                                    case 'darkHorizontal':
                                        $aConvertedAStyle['pattern'] = 0x05;
                                        break;
                                    case 'darkTrellis':
                                        $aConvertedAStyle['pattern'] = 0x0A;
                                        break;
                                    case 'darkUp':
                                        $aConvertedAStyle['pattern'] = 0x08;
                                        break;
                                    case 'darkVertical':
                                        $aConvertedAStyle['pattern'] = 0x06;
                                        break;
                                    case 'gray0625':
                                        $aConvertedAStyle['pattern'] = 0x12;
                                        break;
                                    case 'gray125':
                                        $aConvertedAStyle['pattern'] = 0x11;
                                        break;
                                    case 'lightDown':
                                        $aConvertedAStyle['pattern'] = 0x0D;
                                        break;
                                    case 'lightGray':
                                        $aConvertedAStyle['pattern'] = 0x04;
                                        break;
                                    case 'lightGrid':
                                        $aConvertedAStyle['pattern'] = 0x0F;
                                        break;
                                    case 'lightHorizontal':
                                        $aConvertedAStyle['pattern'] = 0x0B;
                                        break;
                                    case 'lightTrellis':
                                        $aConvertedAStyle['pattern'] = 0x10;
                                        break;
                                    case 'lightUp':
                                        $aConvertedAStyle['pattern'] = 0x0E;
                                        break;
                                    case 'lightVertical':
                                        $aConvertedAStyle['pattern'] = 0x0C;
                                        break;
                                    case 'mediumGray':
                                        $aConvertedAStyle['pattern'] = 0x02;
                                        break;
                                }
                                break;
                            case 'startcolor':
                                if (!empty($value2)) foreach ($value2 as $sKey3 => $value3)
                                {
                                    switch($sKey3) {
                                        case 'argb':
                                            $aConvertedAStyle['fg_color'] = $this->convertARGBToColorIndex($value3);
                                            break;
                                        case 'rgb':
                                            $aConvertedAStyle['fg_color'] = $this->convertRGBToColorIndex($value3);
                                            break;
                                        case 'index':
                                            $aConvertedAStyle['fg_color'] = $value3;
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
                                $aConvertedAStyle['num_format'] = $this->_convertNumFormatCodeToIndex($value2);
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
                                        $aConvertedAStyle['text_h_align'] = 0;
                                        break;
                                    case 'left':
                                        $aConvertedAStyle['text_h_align'] = 1;
                                        break;
                                    case 'right':
                                        $aConvertedAStyle['text_h_align'] = 3;
                                        break;
                                    case 'center':
                                        $aConvertedAStyle['text_h_align'] = 2;
                                        break;
                                    case 'centerContinuous':
                                        $aConvertedAStyle['text_h_align'] = 6;
                                        break;
                                    case 'justify':
                                        $aConvertedAStyle['text_h_align'] = 5;
                                        break;
                                    case 'distributed':
                                        $aConvertedAStyle['text_h_align'] = 7;
                                        break;
                                    case 'fill':
                                        $aConvertedAStyle['text_h_align'] = 4;
                                        break;
                                }
                                break;
                            case 'vertical':
                                switch($value2) {
                                    case 'bottom':
                                        $aConvertedAStyle['text_v_align'] = 2;
                                        break;
                                    case 'top':
                                        $aConvertedAStyle['text_v_align'] = 0;
                                        break;
                                    case 'center':
                                        $aConvertedAStyle['text_v_align'] = 1;
                                        break;
                                    case 'justify':
                                        $aConvertedAStyle['text_v_align'] = 3;
                                        break;
                                    case 'distributed':
                                        $aConvertedAStyle['text_v_align'] = 4;
                                        break;
                                }
                                break;
                            case 'wrap':
                                switch($value2) {
                                    case TRUE:
                                        $aConvertedAStyle['text_wrap'] = 1;
                                        break;
                                    case FALSE:
                                        $aConvertedAStyle['text_wrap'] = 0;
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
                                                        $aConvertedAStyle['border_color'] = $this->convertARGBToColorIndex($value4);
                                                        break;
                                                    case 'rgb':
                                                        $aConvertedAStyle['border_color'] = $this->convertRGBToColorIndex($value4);
                                                        break;
                                                    case 'index':
                                                        $aConvertedAStyle['border_color'] = $value4;
                                                        break;
                                                }
                                            }
                                            break;                                    
                                        case 'style':
                                            switch($value3) {
                                                case 'none':
                                                    $aConvertedAStyle['border'] = 0x00;
                                                    break;
                                                case 'dashDot':
                                                    $aConvertedAStyle['border'] = 0x09;
                                                    break;
                                                case 'dashDotDot':
                                                    $aConvertedAStyle['border'] = 0x0B;
                                                    break;
                                                case 'dashed':
                                                    $aConvertedAStyle['border'] = 0x03;
                                                    break;
                                                case 'dotted':
                                                    $aConvertedAStyle['border'] = 0x04;
                                                    break;
                                                case 'double':
                                                    $aConvertedAStyle['border'] = 0x06;
                                                    break;
                                                case 'hair':
                                                    $aConvertedAStyle['border'] = 0x07;
                                                    break;
                                                case 'medium':
                                                    $aConvertedAStyle['border'] = 0x02;
                                                    break;
                                                case 'mediumDashDot':
                                                    $aConvertedAStyle['border'] = 0x0A;
                                                    break;
                                                case 'mediumDashDotDot':
                                                    $aConvertedAStyle['border'] = 0x0C;
                                                    break;
                                                case 'mediumDashed':
                                                    $aConvertedAStyle['border'] = 0x08;
                                                    break;
                                                case 'slantDashDot':
                                                    $aConvertedAStyle['border'] = 0x0D;
                                                    break;
                                                case 'thick':
                                                    $aConvertedAStyle['border'] = 0x05;
                                                    break;
                                                case 'thin':
                                                    $aConvertedAStyle['border'] = 0x01;
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
        
        return $aConvertedAStyle;
    }
    
    /*
     * Adds a style to the Excel workbook
     * 
     * @param $aStyle Style under array form
     * @return writeexcel_format Style of type WriteExcel
     */
    public function addStyleToWorkbook(&$aStyle = array())
    {
        $oStyle = $this->_writeexcel_workbook->addformat($this->_convertAStyleForWriteExcel($aStyle));
        
        return $oStyle;
    }
    
    /*
     * Closes the Excel workbook saving it
     */
    public function closeWorkbook()
    {
        $this->_writeexcel_workbook->close();
    }
    
    /*
     * Sets the width of one or many columns of an Excel worksheet
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param float $fWidth Width
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstcol Index of the first column
     * @param int $iLastcol Index of the last column
     * @return writeexcel_worksheet|bool Worksheet of type WriteExcel or false
     */
    public function setColumnsWidthToWorksheet($oWorksheet, $fWidth, $iIndexInWorkbook = 0, $iFirstcol = 0, $iLastcol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->set_column($iFirstcol, $iLastcol, $fWidth);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Insertion of an image into an Excel worksheet.  Supports only the type
     * .bmp
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param string $sFilename Chemin du fichier bitmap
     * @param int $iRow Index of the cell line where is inserted the image
     * @param int $iCol Index of the cell column where is inserted the image
     * @return writeexcel_worksheet|bool Worksheet of type WriteExcel or false
     */
    public function insertImageToWorksheet($oWorksheet, $sFilename, $iRow = 0, $iCol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->insert_bitmap($iRow, $iCol, $sFilename);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Permits to calculate and specify the width of a column after its content
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iCol Index of the column
     * @param string $pCellToken Cell content of type NULL, string, int, float, double, bool
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of
     * the columns after the TrueType fonts
     */
    protected function _setColumnAutosizeToWorksheet($oWorksheet, $iCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth && !empty($aStyle))
        {
            if ($iNbCharPercent > -1)
            {
                $sPercentToken = str_pad(',00 %', mb_strlen((string)(int)floor($pCellToken), 'CP1252')+5, '0', STR_PAD_LEFT);
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
                $fCalculatedWidth = mb_strlen((string)$pCellToken, 'CP1252')*1.26;
        }
        
        if (!isset($oWorksheet->_col_sizes[$iCol]) || $fCalculatedWidth > $oWorksheet->_col_sizes[$iCol])
            $oWorksheet->set_column($iCol, $iCol, $fCalculatedWidth);
    }
    
    /*
     * Permits to calculate and specify the width of the columns of cells merged 
     * horizontally depending on the content of the upper left cell
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iFirstCol Index of the first column of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param $pCellToken Value of upper left cell of type NULL, string, int, float, double, bool
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of
     * the columns after the TrueType fonts
     */
    protected function _setMergedColumnsAutosizeToWorksheet($oWorksheet, $iFirstCol = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth && !empty($aStyle))
        {
            if ($iNbCharPercent > -1)
            {
                $sPercentToken = str_pad(',00 %', mb_strlen((string)(int)floor($pCellToken), 'CP1252')+5, '0', STR_PAD_LEFT);
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
                $fCalculatedWidth = mb_strlen((string)$pCellToken, 'CP1252')*1.26;
        }
        
        $iNbCol = $iLastCol - $iFirstCol + 1;
        $fSumColWidth = 0.0;
        
        for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
            $fSumColWidth += $oWorksheet->_col_sizes[$iCol];
        }
        if ($fCalculatedWidth > $fSumColWidth)
        {
            $fRestWidthPerCol = ($fCalculatedWidth - $fSumColWidth) / $iNbCol;
            for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                $oWorksheet->set_column($iCol, $iCol, $oWorksheet->_col_sizes[$iCol] + $fRestWidthPerCol);
            }
        }
    }
    
    /*
     * Write into a cell range with an unique style. The style is overwritten 
     * either if $oStyle is NULL or not.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line to write
     * @param int $iFirstCol Index of the first column to write
     * @param int $iLastRow Index of the last line of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param array $aCellMatrix Cell matrix of the form $aCellMatrix[$iRow][$iCol] = $pValeur
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param array $aMaxStringLengthPerCol Array indexed by column giving the maximal
     * length of the field of the column
     * @param array $aIndexColCellsWithString Index of the cells containing a
     * string
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeFromArrayToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), $oStyle = NULL, &$aStyle = array(), &$aMaxStringLengthPerCol = array(), &$aIndexColCellsWithString = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (!empty($aCellMatrix))
            {
                if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
                
                foreach ($aCellMatrix as $aCol) {
                    $iCurrentRow = $iFirstRow;
                    if (!empty($aCol)) foreach($aCol as $pToken) {
                        if (is_null($pToken)) $pToken = '';
                        if (is_string($pToken)) $pToken = $this->_sanitizeUTF8($pToken);
                        if (is_int($pToken) || is_float($pToken) || is_bool($pToken)) $pToken = (string)$pToken;
                        
                        $oWorksheet->write($iCurrentRow, $iFirstCol, $pToken, $oStyle);
                        
                        ++$iCurrentRow;
                    }
                    ++$iFirstCol;
                }
                
                foreach ($aMaxStringLengthPerCol as $iCol => $aStringLength)
                {
                    if (!empty($aStringLength) && !empty($aCellMatrix[$iCol])) $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $aStringLength[2], $oStyle, $aStyle, $aStringLength[1]);
                }
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes with a style into a cell of an Excel worksheet. The style is overwritten 
     * either if $oStyle is NULL or not.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pToken Content to write of type NULL, string, int, float, double, bool
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pToken = NULL, $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            if (is_null($pToken)) $pToken = '';
            if (is_string($pToken)) $pToken = $this->_sanitizeUTF8($pToken);
            //is_float is an alias of is_double
            if (is_int($pToken) || is_float($pToken) || is_bool($pToken)) $pToken = (string)$pToken;
            
            $oWorksheet->write($iRow, $iCol, $pToken, $oStyle);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $pToken, $oStyle, $aStyle);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a string with a style into a cell of an Excel worksheet. The style 
     * is overwritten either if $oStyle is NULL or not.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param string $sValue Content to write
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeStringToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $sValue = '', $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write_string($iRow, $iCol, $this->_sanitizeUTF8($sValue), $oStyle);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $sValue, $oStyle, $aStyle);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a number with a style into a cell of an Excel worksheet. The style 
     * is overwritten either if $oStyle is NULL or not.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pNumber Content to write of type string, int, float, double
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeNumberToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pNumber = 0, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            //is_float is an alias of is_double
            if (is_int($pNumber) || is_float($pNumber)) $pNumber = (string)$pNumber;
            
            $oWorksheet->write_number($iRow, $iCol, $pNumber, $oStyle);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $pNumber, $oStyle, $aStyle, $iNbCharPercent);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes an empty cell with a style into an Excel worksheet. The style is
     * overwritten only if the cell is already empty and without style.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeBlankToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            $oWorksheet->write_blank($iRow, $iCol, $oStyle);
            if (!isset($oWorksheet->_col_sizes[$iCol]))
                $oWorksheet->set_column($iCol, $iCol, $this->_fDefaultColSize);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes empty cells into an Excel worksheet. The style is
     * overwritten only if the cells are already empty and without style.
     * 
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function writeBlankToManyCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (isset($oStyle) && !$this->checkStyleClass($oStyle)) return FALSE;
            
            for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                    $oWorksheet->write_blank($iRow, $iCol, $oStyle);
                    if (!isset($oWorksheet->_col_sizes[$iCol]))
                        $oWorksheet->set_column($iCol, $iCol, $this->_fDefaultColSize);
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
     * @param writeexcel_worksheet $oWorksheet Worksheet of type WriteExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * 
     * Only to resize the columns:
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * 
     * @return writeexcel_worksheet|bool $oWorksheet Worksheet of type WriteExcel
     * or false
     */
    public function mergeCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->merge_cells($iFirstRow, $iFirstCol, $iLastRow, $iLastCol);
            
            if (!is_null($pCellToken) && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setMergedColumnsAutosizeToWorksheet($oWorksheet, $iFirstCol, $iLastCol, $pCellToken, $oStyle, $aStyle, $iNbCharPercent);
            
            return $oWorksheet;
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
                case $pNumFormat === '$0.00;($0.00)':
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
                case $pNumFormat === '_-#,##0_-;_-"-"#,##0_-;_-@_-':
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
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @param $pNumFormat Number format of type string or integer for index
     * @return writeexcel_format|bool Style of type WriteExcel or false
     */
    public function setNumFormatToStyle($oStyle, $pNumFormat = '')
    {
        if ($this->checkStyleClass($oStyle))
        {
            $oStyle->set_num_format($this->_convertNumFormatCodeToIndex($pNumFormat));
            
             return $oStyle;
        }
        else
            return FALSE;
    }
    
    /*
     * Specifies the text wrap in an Excel style
     * 
     * @param writeexcel_format $oStyle Style of type WriteExcel
     * @return writeexcel_format|bool Style of type WriteExcel or false
     */
    public function setTextWrapToStyle($oStyle)
    {
        if ($this->checkStyleClass($oStyle))
        {
            $oStyle->set_text_wrap();
            
            return $oStyle;
        }
        else
            return FALSE;
    }
}
