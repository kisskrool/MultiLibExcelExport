<?php

namespace MultiLibExcelExport\Excel\Writer\Facade;

use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterFacade;
use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterStyleFacade;

use MultiLibExcelExport\Cell;
use MultiLibExcelExport\CellMatrix;

class ExcelWriterWorksheetFacade extends ExcelWriterFacade
{
    /*
     * Index of the worksheet in the workbook
     */
    protected $_indexInWorkbook;
    
    /*
     * Number of lines of the Excel worksheet from which we export cells in a manner
     * where they are grouped in tables
     */
    const NB_LINES_SWITCH_ARRAY_MODE = 1000;
    
    /*
     * Export des cellules regroupées en tableaux
     */
    protected $_bArrayMode;
    
    /*
     * Constructor
     * 
     * @param string $sWorksheetName Name of the worksheet
     */
    public function __construct($sWorksheetName = '')//verif string
    {
        $this->_indexInWorkbook = 0;
        
        $this->_bArrayMode = FALSE;
        
        $this->_sWorksheet = self::sanitizeTitle($sWorksheetName);
    }
    
    /*
     * Correction of the title to exclude *:/\?[] characters
     * 
     * @param string $sWorksheetName Name of the workwheet
     * @return $sWorksheetName Corrected worksheet name
     */
    public static function sanitizeTitle($sWorksheetName = '')//verif string
    {
        $aSearch = array('*', ':', '/', '\\', '?', '[', ']');
        $sReplace = '-';
        
        return str_replace($aSearch, $sReplace, $sWorksheetName);
    }
    
    /*
     * Adds this worksheet to the Excel workbook
     * 
     * @param int $_indexInWorkbook Index of the worksheet in the workbook
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function addToWorkBook($_indexInWorkbook = 0)
    {
        if ($this->checkAdapterExists(__CLASS__, __METHOD__))
        {
            $this->_worksheet = $this->getAdapter()->addWorksheetToWorkbook($this->_sWorksheet);
            $this->_indexInWorkbook = $_indexInWorkbook;
                
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Defines this worksheet as active in the Excel workbook. Has to be called
     * after the creation of this worksheet and before the creation of the 
     * following ones.
     */
    public function setAsActiveInWorkbook()
    {
        if ($this->checkAdapterExists(__CLASS__, __METHOD__))
        {
            $this->getAdapter()->setWorksheetAsActiveInWorkbook($this->_worksheet, $this->_indexInWorkbook);
        }
        else
            return FALSE;
    }
    
    /*
     * Checks that this worksheet has been added to the Excel workbook
     * 
     * @param string $sClass Class where is called this method
     * @param string $sMethod Method where is called this method
     * @return bool
     */
    protected function checkWorksheetAddedToWorkbook($sClass, $sMethod)
    {
        try {
            if (isset($this->_worksheet))
            {
                return TRUE;
            }
            else
                throw new \Exception($sClass . ' -> ' . $sMethod . ' - The sheet has first to be added to the workbook by the method addToWorkBook().');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
    
    /*
     * Sets the width of one or many columns of this Excel worksheet
     * 
     * @param float $fWidth Width
     * @param int $iFirstcol Index of the first column
     * @param int $iLastcol Index of the last column
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function setColumnsWidth($fWidth, $iFirstcol = 0, $iLastcol = 0)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $this->getAdapter()->setColumnsWidthToWorksheet($this->_worksheet, $fWidth, $this->_indexInWorkbook, $iFirstcol, $iLastcol);
                
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Inserts an image in this Excel worksheet
     * 
     * @param string $sFilename Path to the image file
     * @param int $iRow Index of the line of the cell where we insert the image
     * @param int $iCol Index of the column of the celle where we insert the image
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function insertImage($sFilename, $iRow = 0, $iCol = 0)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $this->getAdapter()->insertImageToWorksheet($this->_worksheet, $sFilename, $iRow, $iCol);
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Deduces if we need to be in array mode or not. The array mode is used for
     * performance reason.
     * 
     * @param array $aCellMatrix Cell matrix of the form $aCellMatrix[$iRow][$iCol]
     * @return bool $bArrayMode
     */
    public function deduceArrayMode(&$aCellMatrix = array())
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            if (count($aCellMatrix) >= self::NB_LINES_SWITCH_ARRAY_MODE 
                    && ($this->getLibrary() === ExcelWriterFacade::PHPEXCEL ||
                        $this->getLibrary() === ExcelWriterFacade::LIBXL ||
                        $this->getLibrary() === ExcelWriterFacade::SPREADSHEETWRITEEXCEL))
                $bArrayMode = TRUE;
            else
                $bArrayMode = FALSE;
            
            return $bArrayMode;
        }
        else
            return FALSE;
    }
    
    /*
     * Specifies the use of array mode for the Excel worksheet
     * 
     * @param bool $bArrayMode
     * @return bool $bArrayMode
     */
    public function setArrayMode($bArrayMode = FALSE)
    {
        $this->_bArrayMode = $bArrayMode;
        
        return $this->_bArrayMode;
    }
    
    /*
     * Writes a matrix of cells
     * 
     * @param array $oCellMatrix Cell matrix in which the member 'cells' is of
     *  the form ->cells[$iRow][$iCol]
     * @param array $aStyleByContentType Associative array of the cell types with
     * their style of type ExcelWriterStyleFacade. The different types of cells are
     * 'info', 'rowhead', 'colhead', 'link', 'float', 'integer', 'percent'. The key 
     * '_default' is used for the default style.
     * @param bool $bAutomaticArrayMode Deduce and specify the array mode
     * automatically
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function writeCellMatrix($oCellMatrix, &$aStyleByContentType = array(), $bAutomaticArrayMode = TRUE)
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            if (!($oCellMatrix instanceof CellMatrix)) return FALSE;
            if (!is_array($oCellMatrix->cells)) return FALSE;
            
            $start_line = 0;//index of the line in which we write the first cell 
            //of the matrix
            
            if ($bAutomaticArrayMode) {
                $bArrayMode = $this->deduceArrayMode($aCellMatrix);
                $this->setArrayMode($bArrayMode);
            }
            
            $aCellMatrixInverse = array();
            if (!empty($oCellMatrix->cells)) foreach ($oCellMatrix->cells as $key_x => $line) {
                $iTmpKeyX = $key_x;

                if (!empty($line)) foreach ($line as $key_y => $cell) {
                    if (!empty($cell)) {
                        if (!($cell instanceof Cell)) return FALSE;
                        
                        if ($cell->getContentType() == 'image') {
                            $cell_params = $cell->getOtherParams();
                            
                            if (isset($cell_params['image_src'])) {
                                $sImageFilename = $cell_params['image_src'];
                                
                                if (file_exists($sImageFilename))
                                    $this->insertImage($sImageFilename, $iTmpKeyX, $key_y);
                            }
                        } else {
                            if (!isset($aCellMatrixInverse[$key_y]))
                                $aCellMatrixInverse[$key_y] = array();
                            $aCellMatrixInverse[$key_y][$iTmpKeyX] = $cell;
                        }
                    }
                }
            }

            // We begin to fetch the cells to construct the table
            $max_x = 0;
            $aMergeCells = array(); //for storage of parameters needed
            //to merge cells AFTER their writing
            if ($this->_bArrayMode)
                $aMaxStringWithLengthPerCol = array();
            if (!empty($aCellMatrixInverse))
                foreach ($aCellMatrixInverse as $key_y => $col) {
                    $sLastContentType = false;
                    if (!empty($col)) {
                        if ($this->_bArrayMode) {
                            $aColCells = array();

                            $aIndexColCellsWithString = array();
                            $iFirstLineToWrite = -1;
                            if (empty($aColCells[$key_y])) {
                                $aColCells[$key_y] = array();
                                $aIndexColCellsWithString[$key_y] = array();
                            }
                        }
                        $iLastKeyX = -1;

                        end($col);
                        $iMaxColX = key($col);
                        reset($col);

                        $iCounterX = 0; //necessary to index the cells of
                        //$aColCells from 0 instead of $key_x in the case
                        //of Spreadsheet::WriteExcel
                        foreach ($col as $key_x => $cell) {

                            $iCoordX = $key_x;
                            if (!empty($start_line))
                                $iCoordX += $start_line;

                            if (!empty($cell))
                                $sContentType = $cell->getContentType();

                            //If cell is void, or of type empty or if the index of
                            //lines aren't continuous, or if the styles are
                            //different
                            if ($this->_bArrayMode &&
                                    ((($key_x - $iLastKeyX > 1) && $iLastKeyX != -1) || (empty($cell) || 'empty' == $cell->getContentType()) || ($sContentType != $sLastContentType && false !== $sLastContentType &&
                                    !empty($aCellMatrixInverse[$key_y][$key_x - 1]) && 'empty' !== $sLastContentType)
                                    )
                            ) {
                                //Writing of the column in table form
                                if (!empty($aColCells[$key_y])) {
                                    if (isset($aStyleByContentType[$sLastContentType]))
                                        $class = $aStyleByContentType[$sLastContentType];
                                    else if (isset($aStyleByContentType['_default']))
                                        $class = $aStyleByContentType['_default'];
                                    else $class = NULL;

                                    $oColStyle = $class;

                                    $this->writeFromArray($iFirstLineToWrite, $key_y, $iCoordX - 1, $key_y, $aColCells, $oColStyle, $aMaxStringWithLengthPerCol, $aIndexColCellsWithString);
                                    $iCounterX = 0;
                                }

                                $aColCells[$key_y] = array();
                                $aIndexColCellsWithString[$key_y] = array();
                                $iFirstLineToWrite = -1;
                            }

                            // If the cell is void or of "empty" type, we pass to the next
                            if (empty($cell) || 'empty' == $cell->getContentType()) {
                                continue;
                            }

                            if ($this->_bArrayMode && $iFirstLineToWrite == -1)
                                $iFirstLineToWrite = $iCoordX;

                            // If the line number of the current cell is superior to
                            // max_x, we copy the value into max_x
                            if ($iCoordX > $max_x)
                                $max_x = $iCoordX;

                            //Merged cells?
                            $x_merge = 0;
                            $y_merge = 0;
                            if ($cell->getRowspan() > 1)
                                $x_merge = $iCoordX + $cell->getRowspan() - 1;
                            if ($cell->getColspan() > 1)
                                $y_merge = $key_y + $cell->getColspan() - 1;

                            // We assign the right cell style in function of its content
                            if (!$this->_bArrayMode || $key_x === $iMaxColX || !empty($x_merge) || !empty($y_merge)) {
                                if (isset($aStyleByContentType[$sContentType]))
                                    $class = $aStyleByContentType[$sContentType];
                                else if (isset($aStyleByContentType['_default']))
                                    $class = $aStyleByContentType['_default'];
                                else $class = NULL;

                                $oColStyle = $class;
                            }

                            // Writing of the cell
                            $pCellToken = NULL;
                            $iNbCharPercent = -1;
                            if ($sContentType === 'link') {
                                // recovery of the additional parameters of the cell
                                $cell_params = $cell->getOtherParams();
                                
                                if (isset($cell_params['link_title'])) {
                                    $pCellToken = $cell_params['link_title'];
                                    if ($this->_bArrayMode)
                                        $aColCells[$key_y][$iCounterX] = $pCellToken;
                                    else
                                        $this->writeString($iCoordX, $key_y, $pCellToken, $class, !empty($x_merge) || !empty($y_merge));
                                }
                            } else if ($sContentType === 'float') {
                                $pCellToken = (float) $cell->getText();

                                if ($this->_bArrayMode)
                                    $aColCells[$key_y][$iCounterX] = $pCellToken;
                                else
                                    $this->writeNumber($iCoordX, $key_y, $pCellToken, $class, !empty($x_merge) || !empty($y_merge));
                            } else if ($sContentType === 'integer') {
                                $pCellToken = floor((float) $cell->getText());

                                if ($this->_bArrayMode)
                                    $aColCells[$key_y][$iCounterX] = $pCellToken;
                                else
                                    $this->writeNumber($iCoordX, $key_y, $pCellToken, $class, !empty($x_merge) || !empty($y_merge));
                            } else if ($sContentType === 'percent') {
                                $pCellToken = ((float) $cell->getText()) / 100;

                                $iNbCharPercent = strlen((string) (int) floor($pCellToken)) + 5;

                                if ($this->_bArrayMode)
                                    $aColCells[$key_y][$iCounterX] = $pCellToken;
                                else
                                    $this->writeNumber($iCoordX, $key_y, $pCellToken, $class, empty($x_merge) || !empty($y_merge), $iNbCharPercent);
                            } else {
                                if ($sContentType == 'text')
                                    $pCellToken = $cell->getText();
                                else
                                    $pCellToken = strip_tags($cell->getText());
                                if ($this->_bArrayMode) {
                                    $aColCells[$key_y][$iCounterX] = $pCellToken;
                                    $aIndexColCellsWithString[$key_y][] = $iCounterX;
                                } else
                                    $this->writeString($iCoordX, $key_y, $pCellToken, $class, !empty($x_merge) || !empty($y_merge));
                            }

                            //length of the line
                            if ($this->_bArrayMode) {
                                //"0 %" in Excel is displayed "0,00 %" by using the format '0.00%', 
                                //equivalent to (whole part + 5) characters
                                if ($sContentType == 'percent')
                                    $iStringLength = strlen((string) (int) floor($pCellToken)) + 5;
                                else
                                    $iStringLength = mb_strlen((string) $pCellToken, 'UTF-8');

                                if ((!isset($aMaxStringWithLengthPerCol[$key_y]) || ($aMaxStringWithLengthPerCol[$key_y][0] < $iStringLength )) &&
                                        empty($x_merge) && empty($y_merge)//not applicable to merged that have their own 
                                //calculation mode regarding the resizing of columns
                                )
                                    $aMaxStringWithLengthPerCol[$key_y] = array($iStringLength, $iNbCharPercent, $pCellToken, $iCoordX);
                            }

                            //Processing of the eventual merging of cells
                            if (!empty($x_merge) || !empty($y_merge)) {
                                if (empty($x_merge))
                                    $x_merge = $iCoordX;
                                if (empty($y_merge))
                                    $y_merge = $key_y;

                                if ($iCoordX + 1 <= $x_merge)
                                    $this->writeBlankToManyCells($iCoordX + 1, $key_y, $x_merge, $key_y, $class);
                                if ($key_y + 1 <= $y_merge)
                                    $this->writeBlankToManyCells($iCoordX, $key_y + 1, $iCoordX, $y_merge, $class);

                                if (!isset($aMergeCells[$y_merge]))
                                    $aMergeCells[$y_merge] = array(); //the index $y_merge
                                //is used to be able to thereafter do an ascending sorting on colspan in order to
                                //avoid cases where for ex. a merged cell with colspan = 2 causes first an 
                                //increase of its columns and then with colspan = 1 grows
                                //also its column provoking a too important growing for the first
                                //cell

                                if ($this->_bArrayMode)//we store the coordinates of the merged cells, because if not
                                //there is a bug with WriteExcel and Spreadsheet::WriteExcel if we merge
                                //the cells before writing them with writeFromArray
                                    $aMergeCells[$y_merge][] = array($iCoordX, $key_y, $x_merge, $y_merge, $pCellToken, $oColStyle, $iNbCharPercent);
                                else//furthermore for the resizing of the cells it is essentiel to merge the 
                                //cells after having written them, otherwise we can face a case where some
                                //columns are heightened after the merging then written and resized without
                                //merging adding finally useless width
                                    $aMergeCells[$y_merge][] = array($iCoordX, $key_y, $x_merge, $y_merge, $pCellToken, $class, $iNbCharPercent);
                            }

                            if ($this->_bArrayMode)
                                $iLastKeyX = $key_x;

                            $sLastContentType = $sContentType;

                            $iCounterX++;
                        }

                        //Writing of the column in table form
                        if ($this->_bArrayMode) {
                            if (!empty($aColCells[$key_y])) {
                                $this->writeFromArray($iFirstLineToWrite, $key_y, $iCoordX, $key_y, $aColCells, $oColStyle, $aMaxStringWithLengthPerCol, $aIndexColCellsWithString);
                            }
                            unset($aColCells);
                            unset($aIndexColCellsWithString);
                        }
                    }
                }
            if ($this->_bArrayMode)
                unset($aMaxStringWithLengthPerCol);

            //We merge the cells if needed
            if (!empty($aMergeCells)) {
                ksort($aMergeCells);
                foreach ($aMergeCells as $aMergeCellsParameters) {
                    foreach ($aMergeCellsParameters as $aParameters)
                        $this->mergeCells($aParameters[0], $aParameters[1], $aParameters[2], $aParameters[3], $aParameters[4], $aParameters[5], $aParameters[6]);
                }
            }
            unset($aMergeCells);
            
            unset($aCellMatrixInverse);

            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes in a cell range. The style is overwritten or not depending on the 
     * used library.
     *
     * @param int $iFirstRow Index of the first line to write
     * @param int $iFirstCol Index of the first column to write
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param array $aCellMatrix Celle matrix in the form $aCellMatrix[$iRow][$iCol] = $pValue
     * @param ExcelWriterStyleFacade $oStyleFacade Style to apply
     * @param array $aMaxStringWithLengthPerCol Table indexed by column giving
     * the field of the column of maximal size with its length
     * @param array $aIndexColCellsWithString Index of the cells containing a string
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function writeFromArray($iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), ExcelWriterStyleFacade $oStyleFacade = NULL, &$aMaxStringWithLengthPerCol = array(), &$aIndexColCellsWithString = array())//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeFromArrayToWorksheet($this->_worksheet, 
                                                           $this->_indexInWorkbook,
                                                           $iFirstRow, 
                                                           $iFirstCol, 
                                                           $iLastRow,
                                                           $iLastCol,
                                                           $aCellMatrix, 
                                                           $oStyle,
                                                           $aStyle,
                                                           $aMaxStringWithLengthPerCol,
                                                           $aIndexColCellsWithString);
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes with a style in a cell of this Excel worksheet. The style is
     * overwritten or not depending on the used library.
     * 
     * @param int $iRow Index of the cell line where we write
     * @param int $iCol Index of the cell column where we write
     * @param $pToken Content to write of type NULL, string, int, float, bool
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function write($iRow = 0, $iCol = 0, $pToken = NULL, ExcelWriterStyleFacade $oStyleFacade = NULL, $bMergedCell = FALSE)//verif int, string
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iRow, $iCol, $pToken, 
                                                  $oStyle,
                                                  $aStyle,
                                                  $bMergedCell
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Write a character string with style in a cell of this Excel worksheet.
     * The style is overwritten or not depending on the used library.
     * 
     * @param int $iRow Index of the cell line where we write
     * @param int $iCol Index of the cell column where we write
     * @param string $sValue Content to write
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @param bool $bMergedCell Indicates if the cell is merged
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function writeString($iRow = 0, $iCol = 0, $sValue = '', ExcelWriterStyleFacade $oStyleFacade = NULL, $bMergedCell = FALSE)//verif int, string
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeStringToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iRow, $iCol, $sValue, 
                                                  $oStyle,
                                                  $aStyle,
                                                  $bMergedCell
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a number with style in a cell of this Excel worksheet. The style
     * is overwritten or not depending on the used library.
     * 
     * @param int $iRow Index of the celle line where we write
     * @param int $iCol Index of the cell column where we write
     * @param $pNumber Conteent to write of type string, int, float, double
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @param bool $bMergedCell Indicates if a cell is merged
     * @param int $iNbCharPercent Number of characters of the field in percentage format
     * @return ExcelWriterWorksheetFacade Cette feuille
     */
    public function writeNumber($iRow = 0, $iCol = 0, $pNumber = 0, ExcelWriterStyleFacade $oStyleFacade = NULL, $bMergedCell = FALSE, $iNbCharPercent = -1)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeNumberToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iRow, $iCol, $pNumber, 
                                                  $oStyle,
                                                  $aStyle,
                                                  $iNbCharPercent,
                                                  $bMergedCell
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Writes a void cell with style in this Excel worksheet. The style is
     * overwritten or not depending on the used library.
     * 
     * @param int $iRow Index of the cell line
     * @param int $iCol Index of the column of the cell
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function writeBlank($iRow = 0, $iCol = 0, ExcelWriterStyleFacade $oStyleFacade = NULL)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeBlankToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iRow, $iCol, 
                                                  $oStyle,
                                                  $aStyle
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Write void cells with style in this Excel worksheet. The style is
     * overwritten or not depending on the used library.
     * 
     * @param int $iFirstRow Index of the first line of the cells
     * @param int $iFirstCol Index of the first column of the cells
     * @param int $iLastRow Index of the last line of the cells
     * @param int $iLastCol Index of the last column of the cells
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function writeBlankToManyCells($iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, ExcelWriterStyleFacade $oStyleFacade = NULL)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->writeBlankToManyCellsToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iFirstRow, $iFirstCol, $iLastRow, $iLastCol,
                                                  $oStyle,
                                                  $aStyle
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Sets a style for many cells in this Excel worksheet. This style is
     * overwrittent or not depending on the used library.
     * 
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function setStyleToManyCells($iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, ExcelWriterStyleFacade $oStyleFacade = NULL)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->setStyleToManyCellsToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iFirstRow, $iFirstCol, $iLastRow, $iLastCol,
                                                  $oStyle,
                                                  $aStyle
                                                  );
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Merges cells in this Excel worksheet
     * 
     * @param int $iFirstRow Index of the first line of cells
     * @param int $iFirstCol Index of the first column of cells
     * @param int $iLastRow Index of the last line of cells
     * @param int $iLastCol Index of the last column of cells
     * 
     * Only to resize the columns:
     * @param $pCellToken Value of the upper left cell of type NULL, string, int, float, double, bool
     * @param ExcelWriterStyleFacade $oStyleFacade Style
     * @param int $iNbCharPercent Number of characters of the field in percetage format
     * 
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function mergeCells($iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, ExcelWriterStyleFacade $oStyleFacade = NULL, $iNbCharPercent = -1)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $oStyle = NULL;
            $aStyle = array();
            if (isset($oStyleFacade))
            {
                if (isset($oStyleFacade->_style)) $oStyle = $oStyleFacade->_style;
                if (isset($oStyleFacade->_aStyle)) $aStyle = $oStyleFacade->_aStyle;
            }
            
            $this->getAdapter()->mergeCellsToWorksheet($this->_worksheet, $this->_indexInWorkbook, $iFirstRow, $iFirstCol, $iLastRow, $iLastCol, $pCellToken, $oStyle, $aStyle, $iNbCharPercent);
            
            return $this;
        }
        else
            return FALSE;
    }
    
    /*
     * Removes a column from an Excel worksheet
     * 
     * @param int $iCol Index of the column to remove
     * @return ExcelWriterWorksheetFacade This worksheet
     */
    public function deleteColumn($iCol = 0)//verif int
    {
        if ($this->checkWorksheetAddedToWorkbook(__CLASS__, __METHOD__) 
                && ($this->getLibrary() == ExcelWriterFacade::LIBXL || 
                    $this->getLibrary() == ExcelWriterFacade::PHPEXCEL))
        {
            $this->getAdapter()->deleteColumnInWorksheet($this->_worksheet, $this->_indexInWorkbook, $iCol);
            
            return $this;
        }
        else
            return FALSE;
    }
            
}