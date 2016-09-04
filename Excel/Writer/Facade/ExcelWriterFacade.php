<?php

namespace MultiLibExcelExport\Excel\Writer\Facade;

class ExcelWriterFacade
{
    const WRITEEXCEL = 'WriteExcel';
    const PHPEXCEL = 'PHPExcel';
    const LIBXL = 'LibXL';
    const SPREADSHEETWRITEEXCEL = 'SpreadsheetWriteExcel';
    
    const PATH_ADAPTER = '/../../Adapter/';
    
    const WRITEEXCEL_ADAPTER = 'Excel_WriteExcel_WorkbookAdapter';
    const PHPEXCEL_ADAPTER = 'Excel_PHPExcel_WorkbookAdapter';
    const LIBXL_ADAPTER = 'Excel_LibXL_WorkbookAdapter';
    const SPREADSHEETWRITEEXCEL_ADAPTER = 'Excel_SpreadsheetWriteExcel_WorkbookAdapter';
    
    /*
     * Static variable initialized in the constructor. Each time an adapter is added
     * it is needed to modify the constructor.
     */
    static protected $_LIBRARIES;
    
    private $_sLibrary;
    private $_oAdapter;
    
    protected $_sWorksheet;
    protected $_worksheet;
    protected $_aStyle;
    protected $_style;
    
    /*
     * Constants defining colors in ARGB, RGB and ColorIndex formats
     */
    const ARGB_BLACK = '00000000';
    const ARGB_WHITE = '00FFFFFF';
    const ARGB_RED = '00FF0000';
    const ARGB_LIME = '0000FF00';
    const ARGB_BLUE = '000000FF';
    const ARGB_YELLOW = '00FFFF00';
    const ARGB_MAGENTA = '00FF00FF';
    const ARGB_CYAN = '0000FFFF';
    const ARGB_BROWN = '00800000';
    const ARGB_GREEN = '00008000';
    const ARGB_NAVY = '00000080';
    const ARGB_PURPLE = '00800080';
    const ARGB_SILVER = '00C0C0C0';
    const ARGB_GRAY = '00808080';
    const ARGB_ORANGE = '00FF6600';
    
    const RGB_BLACK = '000000';
    const RGB_WHITE = 'FFFFFF';
    const RGB_RED = 'FF0000';
    const RGB_LIME = '00FF00';
    const RGB_BLUE = '0000FF';
    const RGB_YELLOW = 'FFFF00';
    const RGB_MAGENTA = 'FF00FF';
    const RGB_CYAN = '00FFFF';
    const RGB_BROWN = '800000';
    const RGB_GREEN = '008000';
    const RGB_NAVY = '000080';
    const RGB_PURPLE = '800080';
    const RGB_SILVER = 'C0C0C0';
    const RGB_GRAY = '808080';
    const RGB_ORANGE = 'FF6600';
    
    const COLINDEX_BLACK = 0x08;
    const COLINDEX_WHITE = 0x09;
    const COLINDEX_RED = 0x0A;
    const COLINDEX_LIME = 0x0B;
    const COLINDEX_BLUE = 0x0C;
    const COLINDEX_YELLOW = 0x0D;
    const COLINDEX_MAGENTA = 0x0E;
    const COLINDEX_CYAN = 0x0F;
    const COLINDEX_BROWN = 0x10;
    const COLINDEX_GREEN = 0x11;
    const COLINDEX_NAVY = 0x12;
    const COLINDEX_PURPLE = 0x14;
    const COLINDEX_SILVER = 0x16;
    const COLINDEX_GRAY = 0x17;
    const COLINDEX_ORANGE = 0x35;
    
    /* 
     * Constructor
     */
    public function __construct()
    {
        self::$_LIBRARIES = array(self::WRITEEXCEL => array('adapter' => self::WRITEEXCEL_ADAPTER), 
                                  self::PHPEXCEL => array('adapter' => self::PHPEXCEL_ADAPTER),
                                  self::LIBXL => array('adapter' => self::LIBXL_ADAPTER),
                                  self::SPREADSHEETWRITEEXCEL => array('adapter' => self::SPREADSHEETWRITEEXCEL_ADAPTER));
    }
    
    /*
     * Accessor of the name of the Excel library
     * 
     * @return string Le nom de la librairie
     */
    public function getLibrary()
    {
        return $this->_sLibrary;
    }
    
    /*
     * Creation of an Excel workbook and its adapter depending on the used library
     * 
     * @param string $sLibrary Name of the library
     * @param string $sWorkbookFilename Path of the file under which we wish to save
     * the Excel workbook
     * @param string $sLibraryPath Path to the Excel libraries
     * @return ExcelWriterFacade This object
     */
    public function createWorkbookAdapterForLibrary($sLibrary, $sWorkbookFilename, $sLibraryPath)//verif string sLibrary
    {
        try {
            if (array_key_exists($sLibrary, self::$_LIBRARIES))
            {
                $this->_sLibrary = $sLibrary;
                $this->_initWorkbookAdapter($sWorkbookFilename, $sLibraryPath);
                
                return $this;
            }
            else
                throw new \Exception('ExcelWriterFacade -> setAdapterForLibrary - The adapter correponding to the library ' . $sLibrary . 'isn\'t referenced.');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
    
    /*
     * Defines the adapter of the used Excel workbook if it has already been created
     * 
     * @param  Excel_Adapter $oAdapter Excel workbook adapter
     * @return ExcelWriterFacade This object
     */
    public function setWorkbookAdapter($oAdapter = NULL)
    {
        if (is_null($oAdapter))
        {
            $this->_oAdapter = NULL;
            $this->_sLibrary = NULL;
        }
        else
        {
            $bAdapterClassFound = false;
            $sLibrary = NULL;
            
            foreach (self::$_LIBRARIES as $sTmpLibrary => $aLibrary)
            {
                if (is_object($oAdapter) && get_class($oAdapter) == "MultiLibExcelExport\\Excel\\Adapter\\".$aLibrary['adapter'])
                {
                    $bAdapterClassFound = true;
                    $sLibrary = $sTmpLibrary;
                }
            }
            
            try {
                if ($bAdapterClassFound)
                {    
                    $this->_sLibrary = $sLibrary;
                    $this->_oAdapter = $oAdapter;
                    
                    return $this;
                }
                else
                    throw new \Exception('ExcelWriterFacade -> setAdapter - Le type de l\'adaptateur transmis en paramètre n\'est pas référencé.');
            }
            catch (\Exception $e) {
                var_dump( htmlentities($e) );
                return FALSE;
            }
        }
    }
    
    /*
     * Accessor of the Excel workbook adapter
     * 
     * @return Excel_Adapter Adapter of Excel workbook
     */
    public function getAdapter()
    {
        return $this->_oAdapter;
    }
    
    /*
     * Checks that the adapter exists
     * 
     * @param string $sClass Class where is called this method
     * @param string $sMethod Method where is called this method
     * @return bool
     */
    protected function checkAdapterExists($sClass, $sMethod)
    {
        try {
            if (isset($this->_oAdapter))
            {
                return TRUE;
            }
            else
                throw new \Exception($sClass . ' -> ' . $sMethod . ' - An adapter has first to be specified by the method createWorkbookAdapterForLibrary() or setWorkbookAdapter().');
        }
        catch (\Exception $e) {
            var_dump( htmlentities($e) );
            return FALSE;
        }
    }
    
    /*
     * Initialization of the adapter of the Excel workbook
     * 
     * @param string $sWorkbookFilename Path to the file under which we wish to
     * save the Excel workbook
     */
    private function _initWorkbookAdapter($sWorkbookFilename, $sLibraryPath)
    {
        $sAdapterName = self::$_LIBRARIES[$this->_sLibrary]['adapter'];
        
        require_once dirname(__FILE__) . self::PATH_ADAPTER . $sAdapterName . '.php';
        
        $sAdapterClass = "MultiLibExcelExport\\Excel\\Adapter\\$sAdapterName";
        $this->_oAdapter = new $sAdapterClass($sWorkbookFilename, $sLibraryPath);
    }
}