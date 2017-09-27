<?php
        require_once dirname((__FILE__)) . '/Classes/PHPExcel.php'; // 
        require_once dirname((__FILE__)) . '/KG_ExpFormats.php'; // Retrieve Format Function

        function Exp4DtoXls($out_filename, $x_offset = 1, $y_offset = 1, $template = "None", $HeaderArray = array(), $data = array()) {
            //parmtype check
            set_time_limit(600);
            if (!is_string($out_filename) || !is_int($x_offset) || !is_int($y_offset) || !is_string($template)) {
                $Proceed = False;
                exit("Arguments are Invalid");
            }
            else {
                $Proceed = True;
                echo "Arguments are Valid\r\n";
            }

            if ($Proceed == True) { 
                if ($x_offset == 0) { // Col 0 must correspond, adjust to -> col A
                    $x_offset += 1;
                }
    
                if ($y_offset == 0) { // Row 0 must correspond, adjust to - > Row 1
                    $y_offset += 1;
                }
    
                echo "Filename: " . $out_filename . "\r\n";
                echo "X Offset: " . strval($x_offset) . ", Y Offset: " . strval($y_offset) . " -> Data Starts On " . NumToAlpha($x_offset).strval($y_offset) .  "\r\n";
    
                $templateArrays = KandGExpFmtTemplate($template); //retrieve formats
                if ($template == "") {
                    $template = "None";
                }
                echo "Format : " . $template . "\r\n";

                if (count($HeaderArray) == 0) { // if headers were not passed
                    $HeaderArray = $templateArrays[0]; // retrieve 'default' headers
                }
                
                $TemplateCmds = $templateArrays[1]; // retrieve Commands
                $StandardExp = $templateArrays[2]; // Standard Paste
                $objPHPExcel = new PHPExcel(); // Create Worksheet, new PHPExcel object

                if ($StandardExp) {
                    $data = ShiftDataArrays ($data);
                    $HdrFmtArray = $templateArrays[3]; // retrieve header formats
                    array_splice($templateArrays,0,3); //remove header, header formats, formula arrays from Data Format Arrays
                    array_unshift($data, $HeaderArray); //add headers to Data Arrays
    
                    echo "\r\nHeaders: \r\n";
                    var_dump($HeaderArray);
                    // echo "Header Format Arrays: \r\n";
                    // var_dump($HdrFmtArray);
    
                    // echo "Data Arrays: \r\n";
                    // foreach ($data as $array) {
                    //     var_dump($array);
                    // }
            
                    // echo "Format Arrays: \r\n";
                    // foreach ($templateArrays as $array) {
                    //     var_dump($array);
                    // }
                    
                    $source = $data;                                     
                    $objPHPExcel->getActiveSheet()->fromArray($source, null, NumToAlpha($x_offset) . strval($y_offset), true);     // Fill worksheet with Headers Array + Data Arrays
            
                    echo "Formatting Headers: ". NumToAlpha($x_offset).strval($y_offset).':'.NumToAlpha($x_offset+count($HeaderArray)-1).strval($y_offset)."\r\n";
                    $ColIdx = -1;
                    foreach ($HdrFmtArray as $styleArray) {  // Set Header Formats
                    ++$ColIdx;
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset).':'.NumToAlpha($x_offset+$ColIdx). strval($y_offset))->applyFromArray($styleArray); // Set Header Cell Style
    
                    $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset).':'.NumToAlpha($x_offset+$ColIdx). strval($y_offset))->getNumberFormat()->applyFromArray($styleArray); // Set Header Cell Formatting
                    }
                    
                    $dataRowCount = $objPHPExcel->getActiveSheet()->getHighestRow();  # of data rows
                    $ColIdx = -1;
                    foreach ($templateArrays as $styleArray) {   // Set Col Range Formats for Data
                        ++$ColIdx;
                        echo "Formatting Range: " . NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx).$dataRowCount . "\r\n";
                        $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx).$dataRowCount)->applyFromArray($styleArray); // Set Data Col Styles
    
                        $objPHPExcel->getActiveSheet()->getStyle(NumToAlpha($x_offset+$ColIdx).strval($y_offset+1).':'.NumToAlpha($x_offset+$ColIdx).$dataRowCount)->getNumberFormat()->applyFromArray($styleArray); // Set Data Col Formatting
                    }
                }
                // Template-Specific 'Commands'
                echo "\r\nExecuting Template Commands: \r\nCmd: ".$TemplateCmds."\r\n";
                // Every Template Function has same parms, [0] excel obj ref, [1] x offset, [2] y offset, [3] Headers [4], Data
                call_user_func($TemplateCmds, $objPHPExcel, $x_offset, $y_offset, $HeaderArray, $data);  // Needs an Error Handler
    
                // Auto Size all Columns
                $dataRowCount = $objPHPExcel->getActiveSheet()->getHighestRow();  # of data rows
                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols       
    
                // Name worksheet and Save
                // $objPHPExcel->getActiveSheet()->setTitle('WorkSheet1');   
                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->save($out_filename);
                echo "\r\nFile Saved: " .$out_filename;
            }
        }
?>