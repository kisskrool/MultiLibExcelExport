<?php

namespace MultiLibExcelExport;

class CellMatrix {
    /*
     * Title of the Excel worksheet
     */
    protected $_title = '';
    
    /*
     * Cell array of type Cell in format [$iRow][$iCol], lines and columns 
     * are numbered from 0
     */
    public $cells = array();
    
    public function __construct($title = '') {
        $this->_title = $title;
    }
    
    public function getTitle()
    {
        return $this->_title;
    }
}