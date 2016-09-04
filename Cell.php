<?php

namespace MultiLibExcelExport;

class Cell {
    /*
     * Type of content among 'text','info','rowhead', 'colhead', 'link', 'float', 
     * 'integer','percent', 'image'
     * The 'empty' type is also available to designate an empty cell.
     */
    protected $_content_type = 'text';
    
    /*
     * Number of merged columns
     */
    protected $_colspan = 1;
    
    /*
     * Number of merged lines
     */
    protected $_rowspan = 1;
    
    /*
     * Content of the cell in raw format
     */
    protected $_text = '';
    
    /*
     * Other parameters
     * For the content type 'link', it is necessary to add the param
     * 'link_title' containing the link title
     * For the content type 'image', it is necessary to add the param
     * 'image_src' specifying the absolute path of the image ot insert. The supported
     * file types are .bmp, .png (not for php_writeexcel), .jpg (not for php_writeexcel),
     * .gif (nor for php_writeexcel, Spreadsheet::WriteExcel, LibXL)
     */
    protected $_otherparams = array();
    
    public function __construct($sText = '', $sContentType = 'text', $iColspan = 1, $iRowspan = 1, $aOtherParams = array())
    {
        $this->_text = $sText;
        $this->_content_type = $sContentType;
        $this->_colspan = $iColspan;
        $this->_rowspan = $iRowspan;
        $this->_otherparams = $aOtherParams;
    }
    
    public function getContentType()
    {
        return $this->_content_type;
    }
    
    public function getColspan()
    {
        return $this->_colspan;
    }
    
    public function getRowspan()
    {
        return $this->_rowspan;
    }
    
    public function getText()
    {
        return $this->_text;
    }
    
    public function getOtherParams()
    {
        return $this->_otherparams;
    }
}