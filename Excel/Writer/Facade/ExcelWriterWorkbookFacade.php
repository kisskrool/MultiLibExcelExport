<?php

namespace MultiLibExcelExport\Excel\Writer\Facade;

use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterFacade;
use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterWorksheetFacade;

class ExcelWriterWorkbookFacade extends ExcelWriterFacade
{
    /*
     * Index of the next worksheet that will be added to the workbook
     */
    protected $_iNextWorksheetIndex;
    
    /* 
     * Constructor
     */
    public function __construct()
    {
        parent::__construct();
        
        $this->_iNextWorksheetIndex = 0;
    }
    
    /*
     * Adds a workshhet to the Excel workbook
     * 
     * @param string $sWorksheetName Name of the worksheet to add
     * @return ExcelWriterWorksheetFacade Added worksheet
     */
    public function addWorksheet($sWorksheetName = '')//verif string
    {
        $oWorksheetFacade = new ExcelWriterWorksheetFacade($sWorksheetName);
        $oWorksheetFacade->setWorkbookAdapter($this->getAdapter());
        $oWorksheetFacade->addToWorkBook($this->_iNextWorksheetIndex);
        $this->_iNextWorksheetIndex++;
        
        return $oWorksheetFacade;
    }
    
    /*
     * Adds a style to the Excel workbook
     * 
     * @param array $aStyle Style under table form
     * @return  Created style
     */
    public function addStyle(array $aStyle = array())
    {
        $oStyleFacade = new ExcelWriterStyleFacade($aStyle);
        $oStyleFacade->setWorkbookAdapter($this->getAdapter());
        $oStyleFacade->addToWorkbook();
        
        return $oStyleFacade;
    }
    
    /*
     * Closes the Excel workbook saving it
     */
    public function close()
    {
        $this->getAdapter()->closeWorkbook();
    }
}