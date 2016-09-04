<?php

namespace MultiLibExcelExport\Excel\Adapter;

interface Excel_WorkbookAdapterInterface
{
    public function addWorksheetToWorkbook($sWorksheetName = '', $iIndexInWorkBook = 0);
    
    public function setWorksheetAsActiveInWorkbook($oWorksheet, $iIndexInWorkBook);
    
    public function addStyleToWorkbook(&$aStyle = array());
    
    public function closeWorkbook();
    
    public function setColumnsWidthToWorksheet($oWorksheet, $fWidth, $iIndexInWorkbook = 0, $iFirstcol = 0, $iLastcol = 0);
    
    public function insertImageToWorksheet($oWorksheet, $sFilename, $iRow = 0, $iCol = 0);
    
    public function writeFromArrayToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, &$aCellMatrix = array(), $oStyle = NULL, &$aStyle = array(), &$aMaxStringWithLengthPerCol = array(), &$aIndexColCellsWithString = array());
    
    public function writeToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pToken = NULL, $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE);
    
    public function writeStringToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $sValue = '', $oStyle = NULL, &$aStyle = array(), $bMergedCell = FALSE);
    
    public function writeNumberToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $pNumber = 0, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1, $bMergedCell = FALSE);
    
    public function writeBlankToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iRow = 0, $iCol = 0, $oStyle = NULL, &$aStyle = array());
    
    public function writeBlankToManyCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $oStyle = NULL, &$aStyle = array());
    
    public function mergeCellsToWorksheet($oWorksheet, $iIndexInWorkbook = 0, $iFirstRow = 0, $iFirstCol = 0, $iLastRow = 0, $iLastCol = 0, $pCellToken = NULL, $oStyle = NULL, &$aStyle = array(), $iNbCharPercent = -1);
    
    public function setNumFormatToStyle($oStyle, $pNumFormat = '');
    
    public function setTextWrapToStyle($oStyle);
}