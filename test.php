<?php

require_once 'includes.php';

use MultiLibExcelExport\MultiLibExcelExport;
use MultiLibExcelExport\CellMatrix;
use MultiLibExcelExport\Cell;
use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterFacade;

$feuille1 = new CellMatrix('feuille1');
$feuille1->cells = array(0 => array(0 => new Cell('cell1'), 
                                    1 => new Cell('cell2')),
                         1 => array(0 => new Cell('cell3'), 
                                    1 => new Cell('', 'image', 1,1, array('image_src' => dirname(__FILE__) . '\\images\\charts.jpg')))
                                );

$feuille2 = new CellMatrix('feuille2');
$feuille2->cells = array(0 => array(0 => new Cell('cell5'), 
                                    1 => new Cell('cell6')),
                         1 => array(0 => new Cell('cell7'), 
                                    1 => new Cell('cell8'))
                                );

$classeur = array(0 => $feuille1,
                  1 => $feuille2);

 //WRITEEXCEL or PHPEXCEL or LIBXL or SPREADSHEETWRITEEXCEL
MultiLibExcelExport::exportToExcel($classeur, dirname(__FILE__) . '\\exports\\excel.xls', ExcelWriterFacade::WRITEEXCEL);
