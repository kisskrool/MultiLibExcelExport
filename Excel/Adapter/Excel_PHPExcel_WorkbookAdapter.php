<?php

namespace MultiLibExcelExport\Excel\Adapter;

use MultiLibExcelExport\Excel\Adapter\Excel_Adapter;
use MultiLibExcelExport\Excel\Adapter\Excel_WorkbookAdapterInterface;

class Excel_PHPExcel_WorkbookAdapter extends Excel_Adapter implements Excel_WorkbookAdapterInterface
{
    static protected $_INCLUDES;
    
    static protected $_WORKSHEET_CLASS;
    static protected $_STYLE_CLASS;
    
    protected $_oPHPExcel;
    protected $_sWorkbookFilename;
    
    protected $_fDefaultColSize;
    protected $_fDefaultRowHeight;
    
    /*
     * References to style tables by cell
     */
    protected $_aStyleByCell;
    
    /*
     * Index of the active worksheet
     */
    protected $_iIndexActiveWorksheet;
    
    /* 
     * Constructor
     * 
     * @param string $sWorkbookFilename File path under which to save the Excel workbook
     * @param string $sLibraryPath Path to the Excel libraries
     */
    public function __construct($sWorkbookFilename, $sLibraryPath)
    {
        self::$__CLASS__ = __CLASS__;
        self::$_INCLUDES = array('/PHPExcel_1.7.9/PHPExcel.php');
        
        self::$_WORKSHEET_CLASS = 'PHPExcel_Worksheet';
        self::$_STYLE_CLASS = 'PHPExcel_Style';
        
        $this->_fDefaultColSize = parent::DEFAULT_COL_SIZE;
        $this->_fDefaultRowHeight = parent::DEFAULT_ROW_HEIGHT;
        
        $this->_iIndexActiveWorksheet = 0;
        
        parent::__construct($sLibraryPath);
        
        $this->_oPHPExcel = new \PHPExcel();
        $this->_sWorkbookFilename = $sWorkbookFilename;
        //We remove the default worksheet because it is useless, worksheets are
        //added when needed
        $this->_oPHPExcel->removeSheetByIndex(0);
        
        $this->_aStyleByCell = array();
        
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array( 'memoryCacheSize'  => '20MB'
                      );
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
    }
    
    /*
     * Adds a worksheeet to the Excel workbook
     * 
     * @param string $sWorksheetName Name of the worksheeet to add
     * @param int $iWorksheetIndex Index of the worksheet in the workbook
     * @return PHPExcel_Worksheet Worksheet of type PHPExcel
     */
    public function addWorksheetToWorkbook($sWorksheetName = '', $iIndexInWorkBook = 0)
    {
        $oWorksheet = $this->_oPHPExcel->createSheet();
        $oWorksheet->setTitle($sWorksheetName);
        
        return $oWorksheet;
    }
    
    /*
     * Sets the workhsset as active in the Excel workbook
     * 
     * @param PHPExcel_Worksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkBook Index of the worksheet in the workbook
     */
    public function setWorksheetAsActiveInWorkbook($oWorksheet, $iIndexInWorkBook)
    {
        $iIndex = $this->_oPHPExcel->getIndex($oWorksheet);
        $this->_oPHPExcel->setActiveSheetIndex($iIndex);
        $this->_iIndexActiveWorksheet = $iIndex;
    }
    
    /*
     * Adds a style to the Excel workbook
     * 
     * All the properties of PHPExcel listed down below are usable. As the
     * property color => index doesn't exist in PHPExcel, it is converted in
     * ARGB.
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
     * @param $aStyle Style in array format
     * @return PHPExcel_Style Style of type PHPExcel
     */
    public function addStyleToWorkbook(&$aStyle = array())
    {
        $oStyle = new \PHPExcel_Style();
        
        //We cannot specify directly a color by its index, therefore we convert
        //in ARGB when it's possible
        if (isset($aStyle['font']['color']['index']))
        {
            $sARGB = $this->convertColorIndexToARGB($aStyle['font']['color']['index']);
            if ($sARGB)
                $aStyle['font']['color']['argb'] = $sARGB;
            unset($aStyle['font']['color']['index']);
        }
        if (isset($aStyle['fill']['startcolor']['index']))
        {
            $sARGB = $this->convertColorIndexToARGB($aStyle['fill']['startcolor']['index']);
            if ($sARGB)
                $aStyle['fill']['startcolor']['argb'] = $sARGB;
            unset($aStyle['fill']['startcolor']['index']);
        }
        if (isset($aStyle['borders']['allborders']['color']['index']))
        {
            $sARGB = $this->convertColorIndexToARGB($aStyle['borders']['allborders']['color']['index']);
            if ($sARGB)
                $aStyle['borders']['allborders']['color']['argb'] = $sARGB;
            unset($aStyle['borders']['allborders']['color']['index']);
        }
        
        //If the format code of number is an index, we convert it into its string
        //value
        if (isset($aStyle['numberformat']['code']) && is_int($aStyle['numberformat']['code']))
            $aStyle['numberformat']['code'] = $this->_convertNumFormatIndexToString($aStyle['numberformat']['code']);
        
        $oStyle->applyFromArray($aStyle);
        
        return $oStyle;
    }
    
    /*
     * Closes the Excel workbook saving it
     */
    public function closeWorkbook()
    {
        //We specify again the active worksheet becauser PHPExcel cannot activate a worksheet
        //just after creation if others are created after
        $this->_oPHPExcel->setActiveSheetIndex($this->_iIndexActiveWorksheet);
        
        $oWriter = \PHPExcel_IOFactory::createWriter($this->_oPHPExcel, 'Excel5');
        $oWriter->save($this->_sWorkbookFilename);
    }
    
    /*
     * Sets the width of one or many columns of an Excel worksheet
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param float $fWidth Width
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstcol Index of the first column
     * @param int $iLastcol Index of the last column
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function setColumnsWidthToWorksheet($oWorksheet, $fWidth, $iIndexInWorkbook = 0, $iFirstcol = 0, $iLastcol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            for ($iColumn = $iFirstcol; $iColumn <= $iLastcol; $iColumn++)
                $oWorksheet->getColumnDimensionByColumn($iColumn)->setWidth($fWidth);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Insertion of an image into an Excel worksheet. Supports the types .bmp, 
     * .png, .jpg, .gif
     * 
     * @param PHPExcel_Worksheet $oWorksheet Workhsheet of type PHPExcel
     * @param string $sFilename File path of the image
     * @param int $iRow Index of the cell line where is inserted the image
     * @param int $iCol Index of the cell column where is inserted the image
     * @return PHPExcel_Worksheet|bool Workhsheet of type PHPExcel or false
     */
    public function insertImageToWorksheet($oWorksheet, $sFilename, $iRow = 0, $iCol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oDrawing = new \PHPExcel_Worksheet_Drawing();
            $oDrawing->setPath($sFilename);
            $oDrawing->setCoordinates(\PHPExcel_Cell::stringFromColumnIndex($iCol).(string)($iRow+1));
            $oDrawing->setWorksheet($oWorksheet);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Calculates the difference between two arrays in a recursive manner
     * comparing keys and values
     * 
     * @param array $aArray1 First array
     * @param array $aArray2 Second array
     * @return array Difference of the two arrays
     */
    protected function _arrayRecursiveDiff($aArray1, $aArray2)
    {
        $aReturn = array();

        foreach ($aArray1 as $mKey => $mValue) 
        {
            if (array_key_exists($mKey, $aArray2)) 
            {
                if (is_array($mValue)) 
                {
                    $aRecursiveDiff = $this->_arrayRecursiveDiff($mValue, $aArray2[$mKey]);
                    if (count($aRecursiveDiff)) $aReturn[$mKey] = $aRecursiveDiff;
                } 
                elseif ($mValue != $aArray2[$mKey]) $aReturn[$mKey] = $mValue;
            }
            else
                $aReturn[$mKey] = $mValue;
        }
        return $aReturn;
    }
    
    /*
     * Permits to calculate and specify the width of a column after the content of 
     * a cell, uses the TrueType fonts available in the folder 
     * PATH_TRUETYPE_FONTS, they have to be added when needed (Arial is used by default)
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iCol Index of the cell colyumn that we merge or in which we
     * write
     * @param int $iRow Index of the cell line that we merge or in which we
     * write
     * @param bool $bCalculateExactColumnWidth Excact calculation of the column 
     * widths after the TrueType fonts
     */
    protected function _setColumnAutosizeToWorksheet($oWorksheet, $iCol = 0, $iRow = 0, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth)
            \PHPExcel_Shared_Font::setAutoSizeMethod(\PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
        else
            \PHPExcel_Shared_Font::setAutoSizeMethod(\PHPExcel_Shared_Font::AUTOSIZE_METHOD_APPROX);
        
        \PHPExcel_Shared_Font::setTrueTypeFontPath(static::$_LIBRARY_PATH . self::PATH_TRUETYPE_FONTS);
        
        $dCurrentWidth = $oWorksheet->getColumnDimensionByColumn($iCol)->getWidth();
        
        $oStyle = $oWorksheet->getStyle(\PHPExcel_Cell::stringFromColumnIndex($iCol) . ($iRow+1));
        $oFont = $oStyle->getFont();
        $oCell = $oWorksheet->getCell(\PHPExcel_Cell::stringFromColumnIndex($iCol) . ($iRow+1));
        $sCellText = $oCell->getFormattedValue();
        
        $fColSize = (float)\PHPExcel_Shared_Font::calculateColumnWidth(
					 $oFont,
                                        $sCellText, 
                                         0,
                                         $oFont);
        
        $fColSize++;
        
        if ($fColSize > $dCurrentWidth)
            $oWorksheet->getColumnDimensionByColumn($iCol)->setWidth($fColSize);
    }
    
    /*
     * Permits to calculate and specify the width of the columns of cells merged 
     * horizontally depending on the content of the upper left cell
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iFirstRow Index of the first line of the cells
     * @param int $iFirstCol Index of the first column of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param bool $bCalculateExactColumnWidth Exact calculation of the width of 
     * the columns according to TrueType fonts
     */
    protected function _setMergedColumnsAutosizeToWorksheet($oWorksheet, $iFirstRow = 0, $iFirstCol = 0, $iLastCol = 0, $bCalculateExactColumnWidth = TRUE)
    {
        if ($bCalculateExactColumnWidth)
            \PHPExcel_Shared_Font::setAutoSizeMethod(\PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
        else
            \PHPExcel_Shared_Font::setAutoSizeMethod(\PHPExcel_Shared_Font::AUTOSIZE_METHOD_APPROX);
        
        \PHPExcel_Shared_Font::setTrueTypeFontPath(static::$_LIBRARY_PATH . self::PATH_TRUETYPE_FONTS);
        
        $iNbCol = $iLastCol - $iFirstCol + 1;
        $dSumColWidth = 0.0;
        
        for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
            $dSumColWidth += $oWorksheet->getColumnDimensionByColumn($iCol)->getWidth();
        }
        
        $oStyle = $oWorksheet->getStyle(\PHPExcel_Cell::stringFromColumnIndex($iFirstCol) . ($iFirstRow+1));
        $oFont = $oStyle->getFont();
        $oCell = $oWorksheet->getCell(\PHPExcel_Cell::stringFromColumnIndex($iFirstCol) . ($iFirstRow+1));
        $sCellText = $oCell->getFormattedValue();
        
        $fCalculatedWidth = (float)\PHPExcel_Shared_Font::calculateColumnWidth(
					 $oFont,
                                        $sCellText,
                                         0,
                                         $oFont);
        $fCalculatedWidth++;
        
        if ($fCalculatedWidth > $dSumColWidth)
        {
            $fRestWidthPerCol = ($fCalculatedWidth - $dSumColWidth) / $iNbCol;
            for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                $dCurrentWidth = $oWorksheet->getColumnDimensionByColumn($iCol)->getWidth();
                $oWorksheet->getColumnDimensionByColumn($iCol)->setWidth($dCurrentWidth + $fRestWidthPerCol);
            }
        }
    }
    
    /*
     * Write into a cell range without overwriting the style. The style is 
     * overwritten wether $aStyle is given or not.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line to write
     * @param int $iFirstCol Index of the first column to write
     * @param int $iLastRow Index of the last line of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param array $aCellMatrix Cell matrix of the form $aCellMatrix[$iRow][$iCol] = $pValue
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @param array $aMaxStringLengthPerCol Array indexed by column giving the maximal
     * length of the field of the column
     * @param array $aIndexColCellsWithString Index of the cells containing a
     * string
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeFromArrayToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), $oStyle = NULL, &$aStyle = array(), &$aMaxStringLengthPerCol = array(), &$aIndexColCellsWithString = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            if (!empty($aCellMatrix))
            {
                $this->setStyleToManyCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow, $iFirstCol, $iLastRow, $iLastCol, $oStyle, $aStyle);
                
                $aCellMatrixInverted = array();
                foreach ($aCellMatrix as $key_y => $aCol) {
                    if (!empty($aCol)) foreach ($aCol as $key_x => $cell)
                    {
                        if (!isset($aCellMatrixInverted)) $aCellMatrixInverted[$key_x] = array();
                        $aCellMatrixInverted[$key_x][$key_y] = $cell;
                    }
                }
                
                $oWorksheet->fromArray($aCellMatrixInverted, NULL, \PHPExcel_Cell::stringFromColumnIndex($iFirstCol).(string)($iFirstRow+1), TRUE);
                unset($aCellMatrixInverted);
                
                foreach ($aMaxStringLengthPerCol as $iCol => $aStringLength)
                {
                    if (!empty($aStringLength) && !empty($aCellMatrix[$iCol])) $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $aStringLength[3]);
                }
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes with a style into a cell of an Excel worksheet. Style isn't overwritten 
     * by $aStyle if it was already defines and if it is identical for this cell.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pToken Content to write of type string, int, float, double, bool
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pToken = NULL, $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->getCellByColumnAndRow($iCol, $iRow+1)->setValue($pToken);
            
            if (isset($aStyle))
            {
                $aArrayDiff = array();
                if (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol])) {
                    $aArrayDiff = $this->_arrayRecursiveDiff($aStyle, $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]);
                }
                
                if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) || 
                    (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) && !empty($aArrayDiff)))
                {
                    $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iCol).(string)($iRow+1) );
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook])) $this->_aStyleByCell[$iIndexInWorkbook] = array();
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow])) $this->_aStyleByCell[$iIndexInWorkbook][$iRow] = array();
                    $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol] = $aStyle;
                }
            }
            
            $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $iRow);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a string into a cell of an Excel worksheet. The style isn't overwritten 
     * by $aStyle if it was already defined and if it is identical for this cell.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param string $sValue Content to write
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeStringToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $sValue = '', $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->getCellByColumnAndRow($iCol, $iRow+1)->setValueExplicit($sValue, \PHPExcel_Cell_DataType::TYPE_STRING);
            
            if (isset($aStyle))
            {
                $aArrayDiff = array();
                if (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol])) {
                    $aArrayDiff = $this->_arrayRecursiveDiff($aStyle, $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]);
                }
                
                if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) || 
                    (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) && !empty($aArrayDiff)))
                {
                    $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iCol).(string)($iRow+1) );
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook])) $this->_aStyleByCell[$iIndexInWorkbook] = array();
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow])) $this->_aStyleByCell[$iIndexInWorkbook][$iRow] = array();
                    $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol] = $aStyle;
                }
            }
            
            $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $iRow);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a number into a cell of an Excel worksheet. The style isn't overwritten 
     * by $aStyle if it was already defined and if it is identical for this cell.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param $pNumber Content to write of type string, int, float, double
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characteres of the field in percentage format
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeNumberToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pNumber = 0, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bMergedCell = FALSE)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->getCellByColumnAndRow($iCol, $iRow+1)->setValueExplicit($pNumber, \PHPExcel_Cell_DataType::TYPE_NUMERIC);
            
            if (isset($aStyle))
            {
                $aArrayDiff = array();
                if (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol])) {
                    $aArrayDiff = $this->_arrayRecursiveDiff($aStyle, $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]);
                }
                
                if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) || 
                    (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) && !empty($aArrayDiff)))
                {
                    $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iCol).(string)($iRow+1) );
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook])) $this->_aStyleByCell[$iIndexInWorkbook] = array();
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow])) $this->_aStyleByCell[$iIndexInWorkbook][$iRow] = array();
                    $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol] = $aStyle;
                }
            }
            
            $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            
            if (!$bMergedCell && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setColumnAutosizeToWorksheet($oWorksheet, $iCol, $iRow);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes an empty cell with a style into an Excel worksheet. The style isn't 
     * overwritten by $aStyle if it was already defined and if it is identical for
     * this cell.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iRow Index of the line of the cell to write
     * @param int $iCol Index of the column of the cell to write
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeBlankToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->getCellByColumnAndRow($iCol, $iRow+1)->setValueExplicit(NULL, \PHPExcel_Cell_DataType::TYPE_NULL);
            
            $dColSize = $oWorksheet->getColumnDimensionByColumn($iCol)->getWidth();
            if ($dColSize == -1)
                $oWorksheet->getColumnDimensionByColumn($iCol)->setWidth($this->_fDefaultColSize);
            
            if (isset($aStyle))
            {
                $aArrayDiff = array();
                if (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol])) {
                    $aArrayDiff = $this->_arrayRecursiveDiff($aStyle, $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]);
                }
                
                if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) || 
                    (isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol]) && !empty($aArrayDiff)))
                {
                    $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iCol).(string)($iRow+1) );
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook])) $this->_aStyleByCell[$iIndexInWorkbook] = array();
                    if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow])) $this->_aStyleByCell[$iIndexInWorkbook][$iRow] = array();
                    $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol] = $aStyle;
                }
            }
            
            $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes empty cells with a style into an Excel worksheet. Style is overwritten 
     * either if passed in parameter $aStyle or not.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function writeBlankToManyCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $aCells = array();
            for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                    if (empty($aCells[$iRow])) $aCells[$iRow] = array();
                    $aCells[$iRow][$iCol] = NULL;
                    
                    $dColSize = $oWorksheet->getColumnDimensionByColumn($iCol)->getWidth();
                    if ($dColSize == -1)
                        $oWorksheet->getColumnDimensionByColumn($iCol)->setWidth($this->_fDefaultColSize);
                }
            }
            
            $oWorksheet->fromArray($aCells, NULL, \PHPExcel_Cell::stringFromColumnIndex($iFirstCol).(string)($iFirstRow+1));
            unset($aCells);
           
            $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iFirstCol).(string)($iFirstRow+1) . ':' .
                                                      \PHPExcel_Cell::stringFromColumnIndex($iLastCol).(string)($iLastRow+1));
            
            for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                    if (isset($aStyle))
                    {
                        if (!isset($this->_aStyleByCell[$iIndexInWorkbook])) $this->_aStyleByCell[$iIndexInWorkbook] = array();
                        if (!isset($this->_aStyleByCell[$iIndexInWorkbook][$iRow])) $this->_aStyleByCell[$iIndexInWorkbook][$iRow] = array();
                        $this->_aStyleByCell[$iIndexInWorkbook][$iRow][$iCol] = $aStyle; 
                    }
                }
                
                $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Apply a style to cells within an Excel worksheet. Style is overwritten 
     * either if passed in parameter $aStyle or not.
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $indexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function setStyleToManyCellsToWorksheet($oWorksheet, $indexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array())
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->duplicateStyleArray($aStyle, \PHPExcel_Cell::stringFromColumnIndex($iFirstCol).(string)($iFirstRow+1) . ':' .
                                                      \PHPExcel_Cell::stringFromColumnIndex($iLastCol).(string)($iLastRow+1));
            
            for ($iRow = $iFirstRow; $iRow <= $iLastRow; $iRow++) {
                for ($iCol = $iFirstCol; $iCol <= $iLastCol; $iCol++) {
                    if (isset($aStyle))
                    {
                        if (!isset($this->_aStyleByCell[$indexInWorkbook])) $this->_aStyleByCell[$indexInWorkbook] = array();
                        if (!isset($this->_aStyleByCell[$indexInWorkbook][$iRow])) $this->_aStyleByCell[$indexInWorkbook][$iRow] = array();
                        $this->_aStyleByCell[$indexInWorkbook][$iRow][$iCol] = $aStyle; 
                    }
                }
                
                $oWorksheet->getRowDimension($iRow+1)->setRowHeight($this->_fDefaultRowHeight);
            }
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Merge cells in an Excel worksheet
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * 
     * Only to resize the columns:
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param array $aStyle Non converted style
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * 
     * @return PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function mergeCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->mergeCellsByColumnAndRow($iFirstCol, $iFirstRow+1, $iLastCol, $iLastRow+1);
            
            if (!is_null($pCellToken) && 
                ((!empty($aStyle) && FALSE === $aStyle['alignment']['wrap']) || empty($aStyle))
                )
                $this->_setMergedColumnsAutosizeToWorksheet($oWorksheet, $iFirstRow, $iFirstCol, $iLastCol);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Deletes a column of an Excel worksheet
     * 
     * @param PHPExcel_Worksheet $oWorksheet Worksheet of type PHPExcel
     * @param int $iIndexInWorkbook Index of the workbook worksheet in facade
     * @param int $iCol Index of the column to delete
     * @param PHPExcel_Worksheet|bool Worksheet of type PHPExcel or false
     */
    public function deleteColumnInWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iCol = 0)
    {
        if ($this->checkWorksheetClass($oWorksheet))
        {
            $oWorksheet->removeColumnByIndex($iCol, 1);
            
            return $oWorksheet;
        }
        else
            return FALSE;
    }
    
    /*
     * Converts a default index format code into its string code, 
     * modeled on the functionn PHPExcel_Style_NumberFormat::fillBuiltInFormatCodes,
     * the index from 0x05 to 0x08 and from 0x29 to 0x2b have been modeled on other
     * adapters.
     * 
     * @param int $iNumFormat Format code of type integer for index
     * @return string Strig of format code
     */
    protected function _convertNumFormatIndexToString($iNumFormat = 0)
    {
        if (is_int($iNumFormat))
        {
            switch ($iNumFormat) {
                case 0x00:
                    $sNumberFormat = \PHPExcel_Style_NumberFormat::FORMAT_GENERAL;
                    break;
                case 0x01:
                    $sNumberFormat = '0';
                    break;
                case 0x02:
                    $sNumberFormat = '0.00';
                    break;
                case 0x03:
                    $sNumberFormat = '#,##0';
                    break;
                case 0x04:
                    $sNumberFormat = '#,##0.00';
                    break;
                case 0x05:
                    $sNumberFormat = '0$;(0$)';
                    break;
                case 0x06:
                    $sNumberFormat = '0$;[Red](0$)';
                    break;
                case 0x07:
                    $sNumberFormat = '$0.00;($0.00)';
                    break;
                case 0x08:
                    $sNumberFormat = '$0.00;[Red]($0.00)';
                    break;
                case 0x09:
                    $sNumberFormat = '0%';
                    break;
                case 0x0a:
                    $sNumberFormat = '0.00%';
                    break;
                case 0x0b:
                    $sNumberFormat = '0.00E+00';
                    break;
                case 0x0c:
                    $sNumberFormat = '# ?/?';
                    break;
                case 0x0d:
                    $sNumberFormat = '# ??/??';
                    break;
                case 0x0e:
                    $sNumberFormat = 'mm-dd-yy';
                    break;
                case 0x0f:
                    $sNumberFormat = 'd-mmm-yy';
                    break;
                case 0x10:
                    $sNumberFormat = 'd-mmm';
                    break;
                case 0x11:
                    $sNumberFormat = 'mmm-yy';
                    break;
                case 0x12:
                    $sNumberFormat = 'h:mm AM/PM';
                    break;
                case 0x13:
                    $sNumberFormat = 'h:mm:ss AM/PM';
                    break;
                case 0x14:
                    $sNumberFormat = 'h:mm';
                    break;
                case 0x15:
                    $sNumberFormat = 'h:mm:ss';
                    break;
                case 0x16:
                    $sNumberFormat = 'm/d/yy h:mm';
                    break;
                case 0x25:
                    $sNumberFormat = '#,##0 ;(#,##0)';
                    break;
                case 0x26:
                    $sNumberFormat = '#,##0 ;[Red](#,##0)';
                    break;
                case 0x27:
                    $sNumberFormat = '#,##0.00;(#,##0.00)';
                    break;
                case 0x28:
                    $sNumberFormat = '#,##0.00;[Red](#,##0.00)';
                    break;
                case 0x29:
                    $sNumberFormat = '_-#,##0_-;_-"-"#,##0_-;_-@_-';
                    break;
                case 0x2a:
                    $sNumberFormat = '_-$* #,##0_-;_-$* "-"#,##0_-;_-$* "-"_-;_-@_-';
                    break;
                case 0x2b:
                    $sNumberFormat = '_-#,##0.00_-;_-"-"#,##0.00_-;_-"-"??_-;_-@_-';
                    break;
                case 0x2c:
                    $sNumberFormat = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
                    break;
                case 0x2d:
                    $sNumberFormat = 'mm:ss';
                    break;
                case 0x2e:
                    $sNumberFormat = '[h]:mm:ss';
                    break;
                case 0x2f:
                    $sNumberFormat = 'mmss.0';
                    break;
                case 0x30:
                    $sNumberFormat = '##0.0E+0';
                    break;
                case 0x31:
                    $sNumberFormat = '@';
                    break;
                default:
                    $sNumberFormat = (string)$iNumFormat;
                    break;
            }
        }
        else
            $sNumberFormat = (string)$iNumFormat;
            
        return $sNumberFormat;
    }
    
    /*
     * Sets a number format in an Excel style
     * 
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @param $pNumFormat Number format of type string or integer for index
     * @return PHPExcel_Style Style of type PHPExcel
     */
    public function setNumFormatToStyle($oStyle, $pNumFormat = '')
    {
        if ($this->checkStyleClass($oStyle))
        {
            if (is_int($pNumFormat))
                $oStyle->getNumberFormat()->setBuiltInFormatCode($pNumFormat);
            else
                $oStyle->getNumberFormat()->setFormatCode($pNumFormat);//The correspondance
            //between default format codes and their default index is automatically
            //determined by PHPExcel
            
            return $oStyle;
        }
        else
            return FALSE;
    }
    
    /*
     * Specifies the text wrap in an Excel style
     * 
     * @param PHPExcel_Style $oStyle Style of type PHPExcel
     * @return PHPExcel_Style Style of type PHPExcel
     */
    public function setTextWrapToStyle($oStyle)
    {
        if ($this->checkStyleClass($oStyle))
        {
            $oStyle->getAlignment()->setWrapText(TRUE);
            return $oStyle;
        }
        else
            return FALSE;
    }
}