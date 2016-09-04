<?php

namespace MultiLibExcelExport\Excel\Writer\Facade;

use MultiLibExcelExport\Excel\Writer\Facade\ExcelWriterFacade;

class ExcelWriterStyleFacade extends ExcelWriterFacade
{
    /*
     * Constructor
     * 
     * @param array $aStyle Style
     */
    public function __construct(array $aStyle = array())
    {
        $this->_aStyle = $this->combineStyleWithDefault($aStyle);
    }
     
    /*
     * Combines the default style with a particular style we wish to apply
     * 
     * @param array $aStyle Style to aply
     * @return array $aStyle Combined style
     */
    protected function combineStyleWithDefault($aStyle){
        $colors_argb = array(
                        'black'=> ExcelWriterFacade::ARGB_BLACK,
                        'white'=> ExcelWriterFacade::ARGB_WHITE
		       );
        $colors_rgb = array(
                        'black'=> ExcelWriterFacade::RGB_BLACK,
                        'white'=> ExcelWriterFacade::RGB_WHITE
		       );
        $colors_colindex = array(
                        'black'=> ExcelWriterFacade::COLINDEX_BLACK,
                        'white'=> ExcelWriterFacade::COLINDEX_WHITE
		       );
         
        $aDefaultStyle = array('font' => array('name' => 'Arial',
                                               'size' => 10,
                                               'bold' => FALSE,
                                               'italic' => FALSE,
                                               'color' => array('argb' => $colors_argb['black'],
                                                                'rgb' => $colors_rgb['black'],
                                                                'index' => $colors_colindex['black']),
                                               'underline' => 'none'),
                                'fill' => array('type' => 'solid',
                                                'startcolor' => array('argb' => $colors_argb['white'],
                                                                      'rgb' => $colors_rgb['white'],
                                                                      'index' => $colors_colindex['white'])),
                                'numberformat' => array('code' => 'General'),
                                'alignment' => array('horizontal' => 'general',
                                                     'vertical' => 'bottom',
                                                     'wrap' => FALSE),
                                'borders' => array('allborders' => array('color' => array('argb' => $colors_argb['black'],
                                                                                          'rgb' => $colors_rgb['black'],
                                                                                          'index' => $colors_colindex['black']),
                                                                         'style' => 'none'
                                                                         ))
                           );
        
        //We take in account the first color properties defined in $aStyle
        //and we remove the following ones        
        $this->setOnlyOneDefaultColorProperty($aStyle['font']['color'], $aDefaultStyle['font']['color']);
        $this->setOnlyOneDefaultColorProperty($aStyle['fill']['startcolor'], $aDefaultStyle['fill']['startcolor']);
        $this->setOnlyOneDefaultColorProperty($aStyle['borders']['allborders']['color'], $aDefaultStyle['borders']['allborders']['color']);
        
        //We combine $aStyle withe the default style
        $aMergedArray = $this->customArrayMerge($aDefaultStyle, $aStyle);
        
        //If the 'fill' property doesn't differ from the default style, we remove it
        //because otherwise it fills a cell with a white color including its border
        //when we want to display the default border color
        //(si 'borders'->'allborders'->'style' = 'none')
        if ($aMergedArray['fill']['type'] == 'solid' && 
                ((isset($aMergedArray['fill']['startcolor']['argb']) && 
                    $aMergedArray['fill']['startcolor']['argb'] == $colors_argb['white']) ||
                 (isset($aMergedArray['fill']['startcolor']['rgb']) && 
                    $aMergedArray['fill']['startcolor']['rgb'] == $colors_rgb['white']) ||
                 (isset($aMergedArray['fill']['startcolor']['index']) && 
                    $aMergedArray['fill']['startcolor']['index'] == $colors_colindex['white'])
                )
            )
            unset($aMergedArray['fill']);
                
        return $aMergedArray;
    }
    
    /*
     * Allows to use only one default color property according with this or those
     * defined in $aStyle
     * 
     * @param array $aStyleColorProperty Color property of $aStyle
     * @param array $aDefaultStyleColorProperty Color property of $aDefaultStyle
     */
    protected function setOnlyOneDefaultColorProperty(&$aStyleColorProperty, &$aDefaultStyleColorProperty)
    {
        if (!empty($aStyleColorProperty))
        {
            $aStyleColorPropertyKeys = array_keys($aStyleColorProperty);
            //We take only the first found property
            $sStyleColorPropertyKey = $aStyleColorPropertyKeys[0];
            
            foreach ($aDefaultStyleColorProperty as $sKey => $pValue) {
                if ($sKey !== $sStyleColorPropertyKey)
                    unset($aDefaultStyleColorProperty[$sKey]);
            }
        }
        else
        {
            $aStyleColorProperty = array();
            unset($aDefaultStyleColorProperty['rgb']);
            unset($aDefaultStyleColorProperty['index']);
        }
    }
    
    /*
     * Replaces the style properties of array1 with those of array2 if they
     * correspond
     * 
     * @param array $aArray1
     * @param array $aArray2
     * @return array
     */
    protected function customArrayMerge(array $aArray1, array $aArray2)
    {
        foreach ($aArray1 as $key => &$val) {
            if (array_key_exists($key, $aArray2)) {
                if (!is_array($val))
                    $aArray1[$key] = $aArray2[$key];
            else
                $aArray1[$key] = $this->customArrayMerge($aArray1[$key], $aArray2[$key]);
         }
     }
        return $aArray1;
    }
    
    /*
     * Accessor of style under a table form
     * 
     * @return array Style
     */
    public function getAStyle()
    {
        return $this->_aStyle;
    }
    
    /*
     * Accessor of style
     * 
     * @return array Style
     */
    public function getStyle()
    {
        return $this->_style;
    }
     
    /*
     * Adds this style to the Excel workbook
     * 
     * @return ExcelWriterStyleFacade This style
     */
    public function addToWorkbook()
    {
        if ($this->checkAdapterExists(__CLASS__, __METHOD__))
        {
            $this->_style = $this->getAdapter()->addStyleToWorkbook($this->_aStyle);
             
            return $this;
        }
        else
            return FALSE;
    }
     
     /*
      * Checks that this style has been added to the Excel workbook
      * 
      * @param string $sClass Class where is called this method
      * @param string $sMethod Method where is called this method
      * @return bool
      */
     protected function checkStyleAddedToWorkbook($sClass, $sMethod)
     {
         try {
             if (isset($this->_style))
             {
                 return TRUE;
             }
             else
                 throw new \Exception($sClass . ' -> ' . $sMethod . ' - The style has first to be added to the workbook by the method addToWorkBook().');
         }
         catch (\Exception $e) {
             var_dump( htmlentities($e) );
             return FALSE;
         }
    }
     
    /*
     * Defines a number format in this Excel style
     * 
     * @param string $sNumFormat Number format
     * @return ExcelWriterStyleFacade This style
     */
    public function setNumFormat($sNumFormat = '')//verif string
    {
        if ($this->checkStyleAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $this->getAdapter()->setNumFormatToStyle($this->_style, $sNumFormat);
            $this->_aStyle = $this->_aStyle;//We duplicate the table to avoid problems
                    //with the passing of the table by reference in the write... functions
                    //of the adapters and predominately those of Excel_PHPExcel_WorkbookAdapter
                    //where the property $_aStyleByCell is used
            
            $this->_aStyle['numberformat']['code'] = $sNumFormat;
             
            return $this;
        }
        else
            return FALSE;
    }
     
    /*
     * Specify the automatic carriage return in this Excel style
     * 
     * @return ExcelWriterStyleFacade This style
     */
    public function setTextWrap()
    {
        if ($this->checkStyleAddedToWorkbook(__CLASS__, __METHOD__))
        {
            $this->getAdapter()->setTextWrapToStyle($this->_style);
            $this->_aStyle = $this->_aStyle;//We duplicate the table in order to avoid
                    //problems with the passing of the table by reference in the
                    //write... functions of the adapters and predominately those
                    //of Excel_PHPExcel_WorkbookAdapter where the property $_aStyleByCell 
                    //is used
            
            $this->_aStyle['alignment']['wrap'] = TRUE;
            return $this;
        }
        else
            return FALSE;
    }
}