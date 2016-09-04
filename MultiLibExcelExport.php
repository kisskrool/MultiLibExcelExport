<?php

namespace MultiLibExcelExport;

use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterFacade;
use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterWorkbookFacade;

/**
 * Demo class for Excel export
 *
 */
class MultiLibExcelExport {
    /*
     * Excel export library
     */

    private static $_EXCEL_EXPORT_LIBRARY = '';

    /*
     * Styles attributable to an Excel workbook of the form
     * array('info' => ExcelWriterStyleFacade,
     *       'rowhead' => ExcelWriterStyleFacade,
     *       'colhead' => ExcelWriterStyleFacade,
     *       'link' => ExcelWriterStyleFacade,
     *       'float' => ExcelWriterStyleFacade,
     *       'integer' => ExcelWriterStyleFacade,
     *       'percent' => ExcelWriterStyleFacade,
     *       '_default' => ExcelWriterStyleFacade//default style
     *      )
     */
    private static $_STYLES_BY_CONTENT_TYPE = array();

    /**
     * Writes utilized styles into the Excel Workbook
     *
     * @param ExcelWriterWorkbookFacade $oWorkbook Workbook
     * @param bool $bDebug Diplays debug infos like execution time and memory
     * used
     *
     * @return bool TRUE if executed successfully
     */
    private static function addStylesToWorkbook($oWorkbook, $bDebug = FALSE) {
        ######################################################################
        #
        # Table of colors
        #
        ######################################################################

        if ($bDebug)
            $timestart = microtime(true);

        $colors = array(
            'black' => ExcelWriterFacade::ARGB_BLACK,
            'blue' => ExcelWriterFacade::ARGB_BLUE,
            'brown' => ExcelWriterFacade::ARGB_BROWN,
            'cyan' => ExcelWriterFacade::ARGB_CYAN,
            'silver' => ExcelWriterFacade::ARGB_SILVER,
            'gray' => ExcelWriterFacade::ARGB_GRAY,
            'green' => ExcelWriterFacade::ARGB_GREEN,
            'lime' => ExcelWriterFacade::ARGB_LIME,
            'magenta' => ExcelWriterFacade::ARGB_MAGENTA,
            'navy' => ExcelWriterFacade::ARGB_NAVY,
            'orange' => ExcelWriterFacade::ARGB_ORANGE,
            'purple' => ExcelWriterFacade::ARGB_PURPLE,
            'red' => ExcelWriterFacade::ARGB_RED,
            'silver' => ExcelWriterFacade::ARGB_SILVER,
            'white' => ExcelWriterFacade::ARGB_WHITE,
            'yellow' => ExcelWriterFacade::ARGB_YELLOW
        );

        //Center alignment
        $center = $oWorkbook->addStyle(array('alignment' => array('horizontal' => 'center')));

        //Heading of columns
        $title = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center')
                )
        );
        //Column heading
        $heading = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE,
                        'color' => array('argb' => $colors['white'])),
                    'fill' => array('startcolor' => array('argb' => $colors['navy'])),
                    'borders' => array('allborders' => array('style' => 'thin',
                            'color' => array('argb' => $colors['white']))),
                    'alignment' => array('horizontal' => 'center')
                )
        );
        //Row heading
        $heading_row = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin'),
                        'color' => array('argb' => $colors['white'])),
                    'alignment' => array('vertical' => 'center',
                        'horizontal' => 'left')
                )
        );

        $heading_row_gray = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin'),
                        'color' => array('argb' => $colors['white'])),
                    'alignment' => array('vertical' => 'center',
                        'horizontal' => 'left'),
                    'fill' => array('startcolor' => array('argb' => $colors['gray']))
                )
        );

        $heading_row_silver = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin'),
                        'color' => array('argb' => $colors['white'])),
                    'alignment' => array('vertical' => 'center',
                        'horizontal' => 'left'),
                    'fill' => array('startcolor' => array('argb' => $colors['silver']))
                )
        );

        //Data
        $normal = $oWorkbook->addStyle(
                array('fill' => array('type' => 'solid',
                        'startcolor' => array('argb' => $colors['white'])),
                    'borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('vertical' => 'center',
                        'horizontal' => 'left')
                )
        );
        //Normal underlined
        $normal_underline = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE,
                        'underline' => 'single'),
                    'borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center')
                )
        );
        //Centered
        $normal_center = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center')
                )
        );
        //Centered percentage
        $normal_center_percent = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '0.00%')
                )
        );

        $normal_center_percent_silver = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '0.00%'),
                    'fill' => array('startcolor' => array('argb' => $colors['silver']))
                )
        );

        //Centered integer
        $normal_center_integer = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '#,##0')
                )
        );

        $normal_center_integer_gray = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '#,##0'),
                    'fill' => array('startcolor' => array('argb' => $colors['gray']))
                )
        );

        $normal_center_integer_silver = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '#,##0'),
                    'fill' => array('startcolor' => array('argb' => $colors['silver']))
                )
        );

        //Centered float
        $normal_center_float = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '0.00')
                )
        );
        //Centered float 4 decimals
        $normal_center_float4dec = $oWorkbook->addStyle(
                array('borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center'),
                    'numberformat' => array('code' => '0.0000')
                )
        );
        //Data
        $bold_center = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'center')
                )
        );
        //Bold / aligned to the left
        $bold_left = $oWorkbook->addStyle(
                array('font' => array('bold' => TRUE),
                    'borders' => array('allborders' => array('style' => 'thin')),
                    'alignment' => array('horizontal' => 'left')
                )
        );
        //Without border
        $without_border = $oWorkbook->addStyle();

        //attribution des styles selon les types de cellules
        self::$_STYLES_BY_CONTENT_TYPE = array('info' => $bold_left,
            'rowhead' => $heading_row,
            'colhead' => $heading,
            'link' => $normal_center,
            'float' => $normal_center_float,
            'integer' => $normal_center_integer,
            'percent' => $normal_center_percent,
            '_default' => $normal_center
        );

        if ($bDebug) {
            $timeend = microtime(true);
            $time = $timeend - $timestart;
            $page_load_time = number_format($time, 3);
            echo "Execution addStyle en " . $page_load_time . " sec<br>";
            echo "Memory used: " . (memory_get_usage() / 1024) / 1024 . 'Mo<br>';
        }

        return TRUE;
    }

    /**
     * Writes a cell matrix into an Excel workbook worksheet
     * taking care of adding styles of each type of cell into the workbook
     *
     * @param ExcelWriterWorksheetFacade $oWorksheet Worksheet of the workbook
     * @param CellMatrix $oCellMatrix Cell matrix
     * @param bool $bDebug Displays debug infos like execution time and memory
     * used
     *
     * @return bool TRUE if executed with success
     */
    private static function cellMatrixToWorksheet($oWorksheet, $oCellMatrix, $bDebug = FALSE) {
        if ($bDebug)
            $timestart = microtime(true);

        $oWorksheet->writeCellMatrix($oCellMatrix, self::$_STYLES_BY_CONTENT_TYPE);

        if ($bDebug) {
            $timeend = microtime(true);
            $time = $timeend - $timestart;
            $page_load_time = number_format($time, 3);
            echo "Execution écriture du tableau en " . $page_load_time . " sec<br>";
            echo "Memory used: " . (memory_get_usage() / 1024) / 1024 . 'Mo<br>';
            $timestart = microtime(true);
        }

        return TRUE;
    }

    /**
     *
     * Exports an ensemble of cell matrix into an Excel workbook
     *
     * @param mixed $aTableCellMatrix - a table containing one or many cell matrix, 
     *                                  one per worksheet of the form array($k => $oCellMatrix, ...)
     *                                  with $k corresponding to the number of 
     *                                  worksheet in Excel workbook (0, 1, 2, ...)
     * @param string $sFilename Absolute path of the file in which we wish to export
     * @param string $sExportLibrary Name of library used for the export
     * @param string $sLibraryPath Absolute path of the directory where are located export libraries
     * @param bool $bDebug Displays debug infos like execution time and memory used
     * 
     * @return bool TRUE if executed with success
     */
    public static function exportToExcel($aTableCellMatrix = array(), $sFilename, $sExportLibrary = '', $sLibraryPath = '', $bDebug = FALSE) {
        set_time_limit(9999);

        if (!is_array($aTableCellMatrix)) {
            return FALSE;
        }

        if ($sLibraryPath == '') $sLibraryPath = dirname(__FILE__) . '\\lib';
        
        if ($bDebug)
            echo "Memory used before Excel export: " . (memory_get_usage() / 1024) / 1024 . 'Mo<br>';

        if ($sExportLibrary == '' || ($sExportLibrary !== ExcelWriterFacade::WRITEEXCEL &&
                $sExportLibrary !== ExcelWriterFacade::PHPEXCEL &&
                $sExportLibrary !== ExcelWriterFacade::LIBXL &&
                $sExportLibrary !== ExcelWriterFacade::SPREADSHEETWRITEEXCEL))
            $sExportLibrary = ExcelWriterFacade::PHPEXCEL;

        self::$_EXCEL_EXPORT_LIBRARY = $sExportLibrary; //WRITEEXCEL or PHPEXCEL or LIBXL or SPREADSHEETWRITEEXCEL

        $sTmpFilename = tempnam('/tmp', 'exp');

        $oWorkbook = new ExcelWriterWorkbookFacade();
        $oWorkbook->createWorkbookAdapterForLibrary(self::$_EXCEL_EXPORT_LIBRARY, $sTmpFilename, $sLibraryPath);

        $bFirstWorksheet = TRUE;
        if (!empty($aTableCellMatrix)) {
            //We define the styles used in the workbook
            self::addStylesToWorkbook($oWorkbook, $bDebug);

            foreach ($aTableCellMatrix as $k => $oCellMatrix) {
                //Name of the worksheet which cannot exceed 31 UTF-8 characters (Excel limitation)
                $sTitle = $oCellMatrix->getTitle();

                $sWorksheetName = $k . '-' . mb_substr($sTitle, 0, 28, 'UTF-8'); //mb_substr
                //is necessary instead of substr in order to not have sliced characters 
                //and avoid bugs in all libraries
                //We create a new worksheet
                $oWorksheet = $oWorkbook->addWorksheet($sWorksheetName);

                if ($bFirstWorksheet) {
                    $oWorksheet->setAsActiveInWorkbook();
                    $bFirstWorksheet = FALSE;
                }

                //We write the cell matrix $oCellMatrix into the worksheet $oWorksheet
                self::cellMatrixToWorksheet($oWorksheet, $oCellMatrix, $bDebug);
            }
        }

        $oWorkbook->close();

        rename($sTmpFilename, $sFilename);

        if ($bDebug)
            echo "Peak memory used after Excel export: " . (memory_get_peak_usage() / 1024) / 1024 . 'Mo<br>';

        return TRUE;
    }

}
