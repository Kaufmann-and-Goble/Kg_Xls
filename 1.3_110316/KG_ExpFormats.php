<?php
    function KandGExpFmtTemplate($template = "None") {
        // Format Arrays ordered [0] - Header Array, [1] - Template Cmds, [2] - Boolean Standard/NonStandard Export [3] - Any Format Arrays for Headers,[4]+ Any Format Arrays for Values
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
            function Cmds_Template_1($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
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
            return array($HdrAr,$Cmds,True,$HdrStyleAr,$styleAr1,$styleAr2,$styleAr3,$styleAr4);
            break;
        case "kLateInterestResults":
            function kLateInterestResults_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
                $data_c = array(); // consolidated employer data
                $titles = $data[0];
                array_push($data_c, $data[1], $data[2], $data[3]); // 2nd, 3rd, 4th arrays
                $hdrs_c = array_slice($hdrs,0,3);
                $hdrs = array_slice ($hdrs, 3);
                array_splice($data,0,4);

                $data[4] = array_map('DateStrToTimeStamp', $data[4]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates
                $data[6] = array_map('DateStrToTimeStamp', $data[6]);

                $data = ShiftDataArrays($data);
                $ar_emp = $data_c[1];
                $data_c = ShiftDataArrays ($data_c);

                $BoldTitle = array(
                    'font'  => array(
                        'bold'  => true,
                        'size'  => 14));

                $BoldItalic = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
                        ),
                    'font'  => array(
                        'bold'  => true,
                        'italic'  => true));

                $BotBorder = array(
                    'borders'  => array(
                        'bottom'  => array(
                            'style' => PHPExcel_Style_Border::BORDER_THIN)));

                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar('1',$x,$y),$titles[0]); // set 1st title               
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar('1',$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs_c, null, NewChar('B',$x,$y).NewChar('4',$x,$y),true); // consolidated hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('4',$x,$y).':'.NewChar('D',$x,$y).NewChar('4',$x,$y).':')->applyFromArray($BotBorder); // Format Title

                $objPHPExcel->getActiveSheet()->fromArray($data_c, null, NewChar('B',$x,$y).NewChar('5',$x,$y),true); // consolidated emp data
                $consdatastrt = intval(NewChar('5',$x,$y));
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $consdataend = $ttlrow+1;
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($ttlrow+1),'Totals');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($ttlrow+1))->applyFromArray($BoldItalic);
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($ttlrow+1),'=SUM('.NewChar('D',$x,$y).NewChar('5',$x,$y).':'.NewChar('D',$x,$y).strval($ttlrow).')'); // calculate sum formula for cons. amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($ttlrow+1))->applyFromArray($BotBorder); 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y)) + 3;

                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y),$titles[1]); // set 2nd title
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y),true); // main data hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($ttlrow+3),$x,$y))->applyFromArray($BotBorder); 
                $objPHPExcel->getActiveSheet()->fromArray($data, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+4),$x,$y),true); // main data
                $datarow = $ttlrow+4;
        
                foreach ($ar_emp as $emp) {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                    $RangeAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NewChar('C',$x,$y).strval($datarow).':'.NewChar('C',$x,$y).NewChar(strval($lastrow),$x,$y)));
                    $keys = array_keys($RangeAr,$emp); // assuming data is already sorted by emp
                    $low = min($keys); // get rows w/ emp
                    $high = max($keys);

                    $objPHPExcel->getActiveSheet()->insertNewRowBefore($datarow+$high+1,2);
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($datarow+$high+1),'Totals');
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($datarow+$high+1))->applyFromArray($BoldItalic); // Bold Italic
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('D',$x,$y).strval($datarow+$low).':'.NewChar('D',$x,$y).strval($datarow+$high).')'); // set emp late deferr total
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('J',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('J',$x,$y).strval($datarow+$low).':'.NewChar('J',$x,$y).strval($datarow+$high).')'); // set emp amt due total 
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                }

                $dataend = $datarow+$high+1;

                // Formats
                // $ signs and red negative values
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($ttlrow),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-'); // cons data amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('H',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('H',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('I',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');

                // dates
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('E',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('E',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('G',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('F',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(2)."-".NumToAlpha($dataColCount);
                $objPHPExcel->GetActiveSheet()->getColumnDimension('A')->setWidth(30);
                for ($i = 2; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                } 
            }

            $Cmds = 'kLateInterestResults_Template';

            return array($HdrAr,$Cmds,False,$HdrStyleAr);
            break;
        case "kLateInterestResults_Cml":
            function kLateInterestResults_Cml_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {

                $data_c = array(); // consolidated employer data
                $titles = $data[0];
                array_push($data_c, $data[1], $data[2], $data[3]); // 2nd, 3rd, 4th arrays
                $hdrs_c = array_slice($hdrs,0,3);
                $hdrs = array_slice ($hdrs, 3);
                array_splice($data,0,4);

                $data[4] = array_map('DateStrToTimeStamp', $data[4]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates
                $data[6] = array_map('DateStrToTimeStamp', $data[6]);

                $data = ShiftDataArrays($data);
                $ar_emp = $data_c[1];
                $data_c = ShiftDataArrays ($data_c);

                $BoldTitle = array(
                    'font'  => array(
                        'bold'  => true,
                        'size'  => 14));

                $BoldItalic = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
                        ),
                    'font'  => array(
                        'bold'  => true,
                        'italic'  => true));

                $BotBorder = array(
                    'borders'  => array(
                        'bottom'  => array(
                            'style' => PHPExcel_Style_Border::BORDER_THIN)));

                $Horiz_Center = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                        ));

                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar('1',$x,$y),$titles[0]); // set 1st title               
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar('1',$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs_c, null, NewChar('B',$x,$y).NewChar('4',$x,$y),true); // consolidated hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('4',$x,$y).':'.NewChar('D',$x,$y).NewChar('4',$x,$y).':')->applyFromArray($BotBorder); // Format Title

                $objPHPExcel->getActiveSheet()->fromArray($data_c, null, NewChar('B',$x,$y).NewChar('5',$x,$y),true); // consolidated emp data
                $consdatastrt = intval(NewChar('5',$x,$y));
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $consdataend = $ttlrow+1;
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($ttlrow+1),'Totals');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($ttlrow+1))->applyFromArray($BoldItalic); // Bold Italic
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($ttlrow+1),'=SUM('.NewChar('D',$x,$y).NewChar('5',$x,$y).':'.NewChar('D',$x,$y).strval($ttlrow).')'); // calculate sum formula for cons. amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($ttlrow))->applyFromArray($BotBorder); 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y)) + 3;

                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y),$titles[1]); // set 2nd title
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y),true); // main data hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($ttlrow+3),$x,$y))->applyFromArray($BotBorder); 
                $objPHPExcel->getActiveSheet()->fromArray($data, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+4),$x,$y),true); // main data
                $datarow = $ttlrow+4;
        
                foreach ($ar_emp as $emp) {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                    $RangeAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NewChar('C',$x,$y).strval($datarow).':'.NewChar('C',$x,$y).NewChar(strval($lastrow),$x,$y)));
                    $keys = array_keys($RangeAr,$emp); // assuming data is already sorted by emp
                    $low = min($keys); // get rows w/ emp
                    $high = max($keys);

                    $objPHPExcel->getActiveSheet()->insertNewRowBefore($datarow+$high+1,2);
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($datarow+$high+1),'Totals');
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($datarow+$high+1))->applyFromArray($BoldItalic); // Bold Italic
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('D',$x,$y).strval($datarow+$low).':'.NewChar('D',$x,$y).strval($datarow+$high).')'); // set emp late deferr total
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('J',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('J',$x,$y).strval($datarow+$low).':'.NewChar('J',$x,$y).strval($datarow+$high).')'); // set emp amt due total 
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('L',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('L',$x,$y).strval($datarow+$low).':'.NewChar('L',$x,$y).strval($datarow+$high).')'); // set emp amt due total 
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('L',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                }

                $dataend = $datarow+$high+1;

                // Formats
                // $ signs and red negative values
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($ttlrow),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-'); // cons data amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('H',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('H',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('I',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('K',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('K',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('L',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');

                // dates
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('E',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('E',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('G',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('F',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');

                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(2)."-".NumToAlpha($dataColCount);
                PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
                $objPHPExcel->GetActiveSheet()->getColumnDimension('A')->setWidth(30);
                for ($i = 2; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                } 
            }

            $Cmds = 'kLateInterestResults_Cml_Template';
            return array($HdrAr,$Cmds,False,$HdrStyleAr);
            break;

        case "YTD_Contributions":
            $HdrAr = array();
            $HdrStyleAr = array();
            function YTD_Contributions_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
                var_dump($hdrs);
                $data[1] = array_map('DateStrToTimeStamp', $data[1]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates
                array_unshift($data, $hdrs);
                $objPHPExcel->getActiveSheet()->fromArray($source, null, NumToAlpha($x_offset) . strval($y_offset), true); // hdrs + data

                $TopRow = array(
                    'borders'  => array(
                        'bottom'  => array(
                            'style' => PHPExcel_Style_Border::BORDER_THIN
                    )
                ),
                'font'  => array(
                    'color' => array('rgb' => '614126')
                )
                );

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).':'.NewChar('F',$x,$y))->applyFromArray($TopRow); // Format Title
                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(1)."-".NumToAlpha($dataColCount);
                for ($i = 1; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                } 
            }
            $Cmds = 'YTD_Contributions_Template';
            return array($HdrAr,$Cmds,False,$HdrStyleAr);
            break;
        case "Template_3":
            break;
        default: // empty arrays (no headers, any formatting, no 'Cmds')
            $HdrAr = array();
            $HdrStyleAr = array();
            function Default_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(1)."-".NumToAlpha($dataColCount);
                for ($i = 1; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                } 
            }
            $Cmds = 'Default_Template';
            return array($HdrAr,$Cmds,True,$HdrStyleAr);
            break;
        }
    }
?>