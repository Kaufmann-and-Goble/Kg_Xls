<?php
        require_once dirname(dirname((__FILE__))) . '/Classes/PHPExcel.php';

        function Exp4DtoXls($out_filename, $x_offset = 1, $y_offset = 1, $template = "None", $data = array()) {
            //parmtype check
            if (!is_string($out_filename) || !is_int($x_offset) || !is_int($y_offset) || !is_string($template)) {
                $Proceed = False;
                echo "Argument Type(s) Are Invalid";
            }
            else {
                $Proceed = True;
                echo "Argument Types Are Valid" . "<br/>";
            }

            if ($Proceed == True) { 
                if ($x_offset == 0) { // Col 0 must correspond, adjust to -> col A
                    $x_offset += 1;
                }
    
                if ($y_offset == 0) { // Row 0 must correspond, adjust - > Row 1
                    $y_offset += 1;
                }
    
                echo "filename: " . $out_filename . "<br/>";
                echo "x Offset: " . strval($x_offset) . " Letter -> " . NumToAlpha($x_offset). "<br/>";
                echo "y Offset: " . strval($y_offset) . "<br/>";
                echo "Format : " . $template . "<br/>";
    
                $templateArrays = KandGExpFmtTemplate($template); //retrieve formats
                $HeaderArray = $templateArrays[0]; //retrieve headers
                $HdrFmtArray = $templateArrays[1]; //retrieve header formats
                $TemplateCmds = $templateArrays[2]; //retrieve Commands
    
                array_splice($templateArrays,0,3); //remove header, header formats, formula arrays from Data Format Arrays
                array_unshift($data, $HeaderArray); //add headers to Data Arrays
    
                echo "<br/>Headers: <br/>";
                var_dump($HeaderArray);
    
                echo "Header Format Arrays: <br/>";
                var_dump($HdrFmtArray);
    
                echo "Data Arrays: <br/>";
                foreach ($data as $array) {
                    var_dump($array);
                }
            
                echo "Format Arrays: <br/>";
                foreach ($templateArrays as $array) {
                    var_dump($array);
                }
                echo "<br/>";
    
                $source = $data;
                
                $objPHPExcel = new PHPExcel(); // Create Worksheet, new PHPExcel object
    
                // Fill worksheet with Headers Array + Data Arrays
                $objPHPExcel->getActiveSheet()->fromArray($source, null, NumToAlpha($x_offset) . strval($y_offset));
                
                $dataRowCount = $objPHPExcel->getActiveSheet()->getHighestRow();  # of data rows
                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
            
                echo "Formatting Headers: ". NumToAlpha($x_offset).strval($y_offset).':'.NumToAlpha($x_offset+count($HdrFmtArray)).strval($y_offset)."<br/>";
                $ColIdx = -1;
                foreach ($HdrFmtArray as $styleArray) {  // Set Header Formats
                    ++$ColIdx;
    
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset).':'.NumToAlpha($x_offset+$ColIdx). strval($y_offset))->applyFromArray($styleArray); // Set Header Cell Style
    
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset).':'.NumToAlpha($x_offset+$ColIdx). strval($y_offset))->getNumberFormat()->applyFromArray($styleArray); // Set Header Cell Formatting
                }
    
                $ColIdx = -1;
                foreach ($templateArrays as $styleArray) {   // Set Col Range Formats for Data
                    ++$ColIdx;
                    echo "Formatting Range: " . NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx).$dataRowCount . "<br/>";
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx)    .$dataRowCount)->applyFromArray($styleArray); // Set Data Col Styles
    
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx)    .$dataRowCount)->getNumberFormat()->applyFromArray($styleArray); // Set Data Col Formatting
                }
    
                // Template-Specific 'Commands'
                echo "<br/>Executing Template Commands: <br/>Cmd: ".$TemplateCmds."<br/>";
                // Every Template Function has same parms, [0] excel obj ref, [1] x offset, [2] y offset
                call_user_func($TemplateCmds, $objPHPExcel, $x_offset, $y_offset);  // Needs an Error Handler
    
                // Auto Size all Columns
                echo "<br/>Auto Sizing Columns ".NumToAlpha(1)."-".NumToAlpha($dataColCount);
                for ($x = 1; $x <= $dataColCount; $x++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($x))->setAutoSize(true);
                } 
            
                echo "<br/>";
    
                // Name worksheet and Save
                $objPHPExcel->getActiveSheet()->setTitle('WorkSheet1');   
                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->save(dirname(dirname((__FILE__))) . '/Out_Xls'. '/'. $out_filename);
                echo "<br/>Out:" . dirname(dirname((__FILE__))) . '/Out_Xls'. '/'. $out_filename;
            }
        }

        function NumToAlpha($n) {
            $n = $n - 1;
            for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
                $r = chr($n%26 + 0x41) . $r;
            return $r;
        }

        function AlphaToNum($a) {
            $l = strlen($a);
            $n = 0;
            for($i = 0; $i < $l; $i++)
                $n = $n*26 + ord($a[$i]) - 0x40;
            return $n;
        }

        function ApplyOffset ($string, $x = 0 , $y = 0) { //assumes $string is always formatted properly with $strt and $end chars

            if ($x > 0) {
                $x -= 1; 
            }

            if ($y > 0) {
                $y -= 1;
            }

            $strt = '~~';
            $end = '**';
            While (true) {
                $strtpos = strpos($string, $strt); // '~~' pos 
                if ($strtpos == False) {
                    break; // Nothing left to Apply
                }
                else {
                    $charpos = $strtpos + 2; // char pos
                    $endpos = strpos($string, $end); // '**' pos
                    $coord = substr($string, $charpos, $endpos - $charpos);
                    if (is_numeric($coord)) {
                        $coord = strval(intval($coord)+$y) ; 
                    }
                    else {
                        $coord = NumToAlpha(AlphaToNum($coord) + $x);
                    }
                $string = substr_replace(substr_replace($string, $coord, $charpos, $endpos - $charpos), '', $strtpos, 2);
                $endpos = strpos($string, $end); // new '**' pos
                $string = substr_replace($string, '', $endpos, 2);
                }
            }
            return $string;            
        }

        function ShiftDataArrays ($data = array()) { // convert 'col' data arrays to 'row' arrays
            $rows = 0;
            $rowarrays = array();
            foreach($data as $array) { // get largest array
               if(count($array) > $rows) {
                 $rows = count($array);
                }
            }

            for ($i=0; $i < $rows; $i++) { // convert
                $rowarray = array(); 
                foreach($data as $array) {
                    if (($i+1) > count($array)){
                        array_push($rowarray,'');
                    }
                    else {
                        array_push($rowarray,$array[$i]);
                    }
                }
                array_push($rowarrays,$rowarray);
            }
            return $rowarrays;
        }

        function KandGExpFmtTemplate($template = "None") {
            // Format Arrays ordered [0] - Header Array, [1] - Format Arrays for Headers, [2] - Template Cmds, [3]+ All Format Arrays for Values

            switch ($template) {
            case "Template_1":

                $HdrAr = array(
                    'Letter', '#', 'Number', 'Real', 'Formula'
                );

                $HdrStyleAr = array(
                array(
                    'font'  => array(
                    'bold'  => true,
                    'size'  => 11,
                    'name'  => 'Palatino'
                    )
                ),
                array(
                    'font'  => array(
                    'bold'  => true,
                    'size'  => 12,
                    'name'  => 'Palatino'
                    )
                ),
                array(
                    'font'  => array(
                    'bold'  => true,
                    'size'  => 13,
                    'name'  => 'Palatino'
                    )
                ),
                array(
                    'font'  => array(
                    'bold'  => true,
                    'size'  => 14,
                    'name'  => 'Palatino'
                    )
                ),
                array(
                    'font'  => array(
                    'bold'  => true,
                    'size'  => 15,
                    'name'  => 'Palatino'
                    )
                )
                );

                function Cmds_Template_1($objPHPExcel, $x = 1, $y = 1) {
                    $d = "D";
                    $objPHPExcel->getActiveSheet()->setCellValue('B18','=SUM(B2:B17)');
                    $objPHPExcel->getActiveSheet()->setCellValue($d.'18','=SUM('.$d.'2:'.$d.'17)');
                }

                $Cmds = 'Cmds_Template_1';

                $styleAr1 = array(
                    'font'  => array(
                        'bold'  => true
                    )
                );

                $styleAr2 = array(
                    'font'  => array(
                    'bold'  => true,
                    'color' => array('rgb' => 'FF0000'),
                    'size'  => 15,
                    'name'  => 'Palatino'
                    )
                );

                $styleAr3 = array(
                    'font'  => array(
                    'bold'  => true,
                    'color' => array('rgb' => '614126'),
                    'size'  => 15,
                    'name'  => 'Palatino'
                    )
                );

                $styleAr4 = array(
                    'font'  => array(
                    'bold'  => false,
                    'color' => array('rgb' => '614126')
                    ),
                    'code' => PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE
                );
                $StyleArrs = array($HdrAr,$HdrStyleAr,$Cmds,$styleAr1,$styleAr2,$styleAr3,$styleAr4);
                return $StyleArrs;
                break;

            case "Template_2":
                break;
            case "Template_3":
                break;
            default: // empty arrays (no headers, any formatting, no 'Cmds')
                $HdrAr = array();

                $HdrStyleAr = array();

                $Cmds = array();

                $StyleArrs = array($HdrAr,$HdrStyleAr,$Cmds);
                return $StyleArrs;
                break;

        }}

?>