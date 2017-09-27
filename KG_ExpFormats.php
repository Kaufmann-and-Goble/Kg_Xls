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

                $hdrs = MultiHdrAdjust($hdrs);
                $hdrs_c = MultiHdrAdjust($hdrs_c);
                $hdr_offset = $hdrs[0];
                $hdr_offset_c = $hdrs_c[0];

                array_shift($hdrs);
                array_shift($hdrs_c);
                array_splice($data,0,4);

                // $data[4] = array_map('DateStrToTimeStamp', $data[4]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates
                // $data[5] = array_map('DateStrToTimeStamp', $data[5]);
                // $data[6] = array_map('DateStrToTimeStamp', $data[6]);

                // array_multisort($data[2], SORT_ASC, $data[1], SORT_NUMERIC, SORT_ASC, $data[0], SORT_ASC, $data[4], SORT_NUMERIC, SORT_ASC, $data[3],  $data[5], $data[6], $data[7], $data[8], $data[9])
                // by employer name, emp #, pcpt name, and then date paid
                 array_multisort( $data[1], SORT_NUMERIC, SORT_ASC, $data[0], SORT_ASC, $data[4], SORT_NUMERIC, SORT_ASC, $data[2], $data[3],  $data[5], $data[6], $data[7], $data[8], $data[9]);
                 // by emp #, pcpt name, and then date paid

                $data = ShiftDataArrays($data);
                $ar_emp = $data_c[0];
                // array_multisort($data_c[1], SORT_ASC, $data_c[0], $data_c[2]);
                // by emp name

                array_multisort($data_c[0], SORT_NUMERIC, SORT_ASC, $data_c[1], $data_c[2]);
                // by emp #
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
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y).':')->applyFromArray($BotBorder); // Format Title

                $objPHPExcel->getActiveSheet()->fromArray($data_c, null, NewChar('B',$x,$y).NewChar(strval(5+$hdr_offset_c),$x,$y),true); // consolidated emp data
                $consdatastrt = intval(NewChar('5',$x,$y));                                 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($Horiz_Center); // format emp #
                $consdataend = $ttlrow+1;
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($ttlrow+1),'Totals');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($ttlrow+1))->applyFromArray($BoldItalic); // Bold Italic
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($ttlrow+1),'=SUM('.NewChar('D',$x,$y).NewChar(strval(5+$hdr_offset_c),$x,$y).':'.NewChar('D',$x,$y).strval($ttlrow).')'); // calculate sum formula for cons. amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($ttlrow))->applyFromArray($BotBorder); 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y)) + 3;
        
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y),$titles[1]); // set 2nd title
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y),true); // main data hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y))->applyFromArray($BotBorder); 
                $objPHPExcel->getActiveSheet()->fromArray($data, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+4+$hdr_offset),$x,$y),true); // main data
                $datarow = $ttlrow+4+$hdr_offset;                           

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('4',$x,$y).':'.NewChar('D',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y))->applyFromArray($Horiz_Center); // align emp headers
                $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y))->applyFromArray($Horiz_Center); // align data headers
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($ttlrow+4),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Horiz_Center); // align text data 'B'
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('E',$x,$y).NewChar(strval($ttlrow+4),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Horiz_Center); // align text data, 'E-G'

                foreach ($ar_emp as $emp) {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('B',$x,$y));
                    $RangeAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NewChar('B',$x,$y).strval($datarow).':'.NewChar('B',$x,$y).NewChar(strval($lastrow),$x,$y)));
                    $keys = array_keys($RangeAr,$emp); // assuming data is already sorted by emp
                    $low = min($keys); // get rows w/ emp
                    $high = max($keys);

                    $range = $high-$low+1;
                    $high2+=$range+2;

                    $objPHPExcel->getActiveSheet()->insertNewRowBefore($datarow+$high+1,2);
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($datarow+$high+1),'Totals');
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($datarow+$high+1))->applyFromArray($BoldItalic); // Bold Italic
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('D',$x,$y).strval($datarow+$low).':'.NewChar('D',$x,$y).strval($datarow+$high).')'); // set emp late deferr total
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                    $objPHPExcel->getActiveSheet()->setCellValue(NewChar('J',$x,$y).strval($datarow+$high+1),'=SUM('.NewChar('J',$x,$y).strval($datarow+$low).':'.NewChar('J',$x,$y).strval($datarow+$high).')'); // set emp amt due total 
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).strval($datarow+$high))->applyFromArray($BotBorder); // Format Total
                }

                $dataend = $datarow+$high2-1;

                // Formats
                // $ signs and red negative values
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($ttlrow),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-'); // cons data amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('H',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('H',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('I',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');

                for ($char='E'; $char <= 'G'; $char++) { 
                    for ($i=$datarow; $i<($dataend+1) ; $i++) { 
                        $value = $objPHPExcel->getActiveSheet()->getCell(NewChar($char,$x,$y).strval($i))->getValue();
                        if (CheckDateFormat($value)) { // if parsable date format, convert to Excel Date
                            $objPHPExcel->GetActiveSheet()->setCellValue(NewChar($char,$x,$y).strval($i), DateStrToTimeStamp($value));
                            $objPHPExcel->GetActiveSheet()->getStyle(NewChar($char,$x,$y).strval($i))->getNumberFormat()->setFormatCode('mm/dd/y;;'); 
                        } else { // leave as text
                        
                        }                    
                    }
                }

                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(2)."-".NumToAlpha($dataColCount);
                $objPHPExcel->GetActiveSheet()->getColumnDimension('A')->setWidth(30);
                for ($i = 2; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                }
                $objPHPExcel->GetActiveSheet()->setSelectedCell('A1'); 
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

                $hdrs = MultiHdrAdjust($hdrs);
                $hdrs_c = MultiHdrAdjust($hdrs_c);
                $hdr_offset = $hdrs[0];
                $hdr_offset_c = $hdrs_c[0];

                array_shift($hdrs);
                array_shift($hdrs_c);
                array_splice($data,0,4); // titles and consolidated arrays

                // array_multisort($data[2], SORT_ASC, $data[1], SORT_NUMERIC, SORT_ASC, $data[0], SORT_ASC, $data[4], SORT_NUMERIC, SORT_ASC, $data[3], $data[5], $data[6], $data[7], $data[8], $data[9], $data[10], $data[11]);
                // by emp name, emp #, pcpt name, date paid

                array_multisort($data[1], SORT_NUMERIC, SORT_ASC, $data[0], SORT_ASC, $data[4], SORT_NUMERIC, SORT_ASC, $data[2], $data[3], $data[5], $data[6], $data[7], $data[8], $data[9], $data[10], $data[11]);
                // by emp #, pcpt name, date paid

                $data = ShiftDataArrays($data);
                // $ar_emp = $data_c[1];
                $ar_emp = $data_c[0];
                // array_multisort($data_c[1], SORT_ASC, $data_c[0], $data_c[2]);
                // by emp name

                array_multisort($data_c[0], SORT_NUMERIC, SORT_ASC, $data_c[1], $data_c[2]);
                // by emp name

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
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y).':')->applyFromArray($BotBorder); // Format Title

                $objPHPExcel->getActiveSheet()->fromArray($data_c, null, NewChar('B',$x,$y).NewChar(strval(5+$hdr_offset_c),$x,$y),true); // consolidated emp data
                $consdatastrt = intval(NewChar('5',$x,$y));                                 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($Horiz_Center); // format emp #
                $consdataend = $ttlrow+1;
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('C',$x,$y).strval($ttlrow+1),'Totals');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).strval($ttlrow+1))->applyFromArray($BoldItalic); // Bold Italic
                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('D',$x,$y).strval($ttlrow+1),'=SUM('.NewChar('D',$x,$y).NewChar(strval(5+$hdr_offset_c),$x,$y).':'.NewChar('D',$x,$y).strval($ttlrow).')'); // calculate sum formula for cons. amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).strval($ttlrow))->applyFromArray($BotBorder); 
                $ttlrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y)) + 3;

                $objPHPExcel->getActiveSheet()->setCellValue(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y),$titles[1]); // set 2nd title
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow),$x,$y))->applyFromArray($BoldTitle); // Format Title
                $objPHPExcel->getActiveSheet()->fromArray($hdrs, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y),true); // main data hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y))->applyFromArray($BotBorder); 
                $objPHPExcel->getActiveSheet()->fromArray($data, null, NewChar('A',$x,$y).NewChar(strval($ttlrow+4+$hdr_offset),$x,$y),true); // main data
                $datarow = $ttlrow+4+$hdr_offset;                           

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('4',$x,$y).':'.NewChar('D',$x,$y).NewChar(strval(4+$hdr_offset_c),$x,$y))->applyFromArray($Horiz_Center); // align emp headers
                $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('C',$x,$y));
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($ttlrow+3),$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($ttlrow+3+$hdr_offset),$x,$y))->applyFromArray($Horiz_Center); // align data headers
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($ttlrow+4),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Horiz_Center); // align text data 'B'
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('E',$x,$y).NewChar(strval($ttlrow+4),$x,$y).':'.NewChar('G',$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Horiz_Center); // align text data, 'E-G'

                foreach ($ar_emp as $emp) {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('B',$x,$y));
                    $RangeAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NewChar('B',$x,$y).strval($datarow).':'.NewChar('B',$x,$y).NewChar(strval($lastrow),$x,$y)));
                    $keys = array_keys($RangeAr,$emp); // assuming data is already sorted by emp
                    $low = min($keys); // get rows w/ emp
                    $high = max($keys);

                    $range = $high-$low+1;
                    $high2+=$range+2;

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

                $dataend = $datarow+$high2-1;

                // Formats
                // $ signs and red negative values
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($consdatastrt),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($ttlrow),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-'); // cons data amt due
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('H',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('H',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('I',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('J',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('K',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('K',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('L',$x,$y).NewChar(strval($datarow),$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($dataend),$x,$y))->getNumberFormat()->setFormatCode('"$"#,##0.00_-;[Red]-"$"#,##0.00_-;"$"#,##0.00_-');

                for ($char='E'; $char <= 'G'; $char++) { 
                    for ($i=$datarow; $i<($dataend+1) ; $i++) { 
                        $value = $objPHPExcel->getActiveSheet()->getCell(NewChar($char,$x,$y).strval($i))->getValue();
                        if (CheckDateFormat($value)) { // if parsable date format, convert to Excel Date
                            $objPHPExcel->GetActiveSheet()->setCellValue(NewChar($char,$x,$y).strval($i), DateStrToTimeStamp($value));
                            $objPHPExcel->GetActiveSheet()->getStyle(NewChar($char,$x,$y).strval($i))->getNumberFormat()->setFormatCode('mm/dd/y;;'); 
                        } else { // leave as text
                        
                        }                    
                    }
                }

                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(2)."-".NumToAlpha($dataColCount);
                $objPHPExcel->GetActiveSheet()->getColumnDimension('A')->setWidth(30);
                for ($i = 2; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                }
                $objPHPExcel->GetActiveSheet()->setSelectedCell('A1'); 
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
        case "BadCensusAdds":
            function BadCensusAdds_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
                $hdrs2 = $hdrs; // copy and remove last 2 headers 
                $data[8] = array_map('DateStrToTimeStamp', $data[8]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates
                $data[11] = array_map('DateStrToTimeStamp', $data[11]);               
                $data = ShiftDataArrays($data);
                $adj_hdrs = array();
                $adj_hdrs = MultiHdrAdjust($hdrs);
                $hdr_offset = $adj_hdrs[0];
                array_shift($adj_hdrs);
                // var_dump($adj_hdrs);

                $BoldCenter = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                        ),
                    'font'  => array(
                        'size'  => 12,
                        'bold'  => true
                        ));

                $BoldLeft = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
                        ),
                    'font'  => array(
                        'size'  => 12,
                        'bold'  => true
                        ));

                $CenterSize12 = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                        ),
                    'font'  => array(
                        'size'  => 12
                        ));

                $LeftSize12 = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
                        ),
                    'font'  => array(
                        'size'  => 12
                        ));

                $Size12 = array(
                    'font'  => array(
                    'size'  => 12
                    ));

                $objPHPExcel->getActiveSheet()->fromArray($adj_hdrs, null, NewChar('A',$x,$y).NewChar('1',$x,$y),true); // hdrs
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar('1',$x,$y).':'.NewChar('A',$x,$y).NewChar('2'),$x,$y)->applyFromArray($BoldCenter);
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('1',$x,$y).':'.NewChar('H',$x,$y).NewChar('2'),$x,$y)->applyFromArray($BoldLeft);
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar('1',$x,$y).':'.NewChar('I',$x,$y).NewChar('2'),$x,$y)->applyFromArray($BoldCenter);
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('J',$x,$y).NewChar('1',$x,$y).':'.NewChar('K',$x,$y).NewChar('2'),$x,$y)->applyFromArray($BoldLeft);
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('L',$x,$y).NewChar('1',$x,$y).':'.NewChar('L',$x,$y).NewChar('2'),$x,$y)->applyFromArray($BoldCenter);

                $objPHPExcel->getActiveSheet()->fromArray($data, null, NewChar('A',$x,$y).NewChar(strval(2+$hdr_offset),$x,$y),true); // data
                $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                for ($char='A'; $char <= 'L'; $char++) { 
                    switch ($char) {
                        case ($char == 'A') || ($char == 'I') || ($char == 'L'):
                            $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar('1',$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($CenterSize12);
                            break;
                        default:
                            $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar('1',$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($LeftSize12);
                            break;
                    }
                }

                $objPHPExcel->getActiveSheet()->getStyle(NewChar('I',$x,$y).NewChar('1',$x,$y).':'.NewChar('I',$x,$y).NewChar(strval($lastrow),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y');
                $objPHPExcel->getActiveSheet()->getStyle(NewChar('L',$x,$y).NewChar('1',$x,$y).':'.NewChar('L',$x,$y).NewChar(strval($lastrow),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/y;;');

                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                // $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                // $RangeAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NewChar('A',$x,$y).NewChar('1',$x,$y).':'.NewChar('A',$x,$y).NewChar(strval($lastrow),$x,$y)));
                
                // if ($high>$high2) {
                //     $high = $high2;
                // } 
                
                // $objPHPExcel->getActiveSheet()->insertNewRowBefore($high+2,2);
                // $objPHPExcel->getActiveSheet()->fromArray($hdrs2, null, NewChar('A',$x,$y).strval($high+2),true); // hdrs

                echo "\r\nAuto Sizing Columns ".NumToAlpha(1)."-".NumToAlpha($dataColCount);
                for ($i = 1; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                }
                $objPHPExcel->GetActiveSheet()->setSelectedCell('A1');  
            }    
            $Cmds = 'BadCensusAdds_Template';
            return array($HdrAr,$Cmds,False,$HdrStyleAr);
            break;

        case "BadCensusDOBs":
            function BadCensusDOBs_Template($objPHPExcel, $x = 1, $y = 1, $hdrs = array(), $data = array()) {
                $hdrs3 = $hdrs;
                $hdrs2 = array_slice($hdrs,0,3);
                array_push($hdrs2, $hdrs[4]);
                $hdrs = array_slice($hdrs, 0, 3);

                $data[10] = array_map('DateStrToTimeStamp', $data[10]); // convert datestring->UNIX timestamp->Excel Serial Number Format for Dates

                $data1 = array();
                $data2 = array();
                $data3 = array();
                array_push($data1, $data[0], $data[1], $data[2]);
                array_push($data2, $data[3], $data[4], $data[5], $data[6]);
                array_push($data3, $data[7], $data[8], $data[9], $data[10], $data[11]);
                $bool1 = True;
                $bool2 = True;
                $bool3 = True;

                if (count($data[0])==1 & $data[0][0]=="") {
                    $bool1 = False;
                }

                if (count($data[3])==1 & $data[3][0]=="") {
                    $bool2 = False;
                }

                if (count($data[7])==1 & $data[7][0]=="") {
                    $bool3 = False;
                }

                $data1 = ShiftDataArrays($data1);
                $data2 = ShiftDataArrays($data2);
                $data3 = ShiftDataArrays($data3);

                $BoldCenter = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                        ),
                    'font'  => array(
                        'size'  => 12,
                        'bold'  => true
                        ));

                $BoldLeft = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
                        ),
                    'font'  => array(
                        'size'  => 12,
                        'bold'  => true
                        ));

                $CenterSize12 = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                        ),
                    'font'  => array(
                        'size'  => 12
                        ));

                $LeftSize12 = array(
                    'alignment'  => array(
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
                        ),
                    'font'  => array(
                        'size'  => 12
                        ));

                $Size12 = array(
                    'font'  => array(
                    'size'  => 12
                    ));
                

                if ($bool1) {
                    $adj_hdrs = array();
                    $adj_hdrs = MultiHdrAdjust($hdrs);
                    $hdr_offset = $adj_hdrs[0];
                    array_shift($adj_hdrs);

                    $objPHPExcel->getActiveSheet()->fromArray($adj_hdrs, null, NewChar('A',$x,$y).NewChar('1',$x,$y),true); // hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar('1',$x,$y).':'.NewChar('A',$x,$y).NewChar('1'),$x,$y)->applyFromArray($BoldLeft); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar('1',$x,$y).':'.NewChar('B',$x,$y).NewChar('1'),$x,$y)->applyFromArray($BoldCenter); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).NewChar('1',$x,$y).':'.NewChar('C',$x,$y).NewChar('1'),$x,$y)->applyFromArray($BoldLeft); // format hdrs

                    $objPHPExcel->getActiveSheet()->fromArray($data1, null, NewChar('A',$x,$y).NewChar(strval(2+$hdr_offset),$x,$y),true); // data
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                    for ($char='A'; $char <= 'C'; $char++) {
                        switch ($char) {
                            case ($char == 'B'):
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($CenterSize12);

                                break;
                            default:
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Size12);
                                break;
                        }
                    }

                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($lastrow),$x,$y))->getNumberFormat()->setFormatCode('000000000');
                
                }
                else {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                }

                if ($bool2) {
                    $adj_hdrs = array();
                    $adj_hdrs = MultiHdrAdjust($hdrs2);
                    $hdr_offset = $adj_hdrs[0];
                    array_shift($adj_hdrs);

                    $objPHPExcel->getActiveSheet()->fromArray($adj_hdrs, null, NewChar('A',$x,$y).NewChar(strval($lastrow+2),$x,$y),true); // hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('A',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldLeft); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldCenter); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('C',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldLeft); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldCenter); // format hdrs

                    $objPHPExcel->getActiveSheet()->fromArray($data2, null, NewChar('A',$x,$y).NewChar(strval($lastrow+3),$x,$y),true); // 2nd data

                    $firstrow = $lastrow+4;
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                    for ($char='A'; $char <= 'D'; $char++) { 
                        switch ($char) {
                            case ($char == 'B') || ($char == 'D') :
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($CenterSize12);

                                break;
                            default:
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Size12);
                                break;
                        }
                    }
                }
                else {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                }

                if ($bool3) {
                    $adj_hdrs = array();
                    $adj_hdrs = MultiHdrAdjust($hdrs3);
                    $hdr_offset = $adj_hdrs[0];
                    array_shift($adj_hdrs);

                    $objPHPExcel->getActiveSheet()->fromArray($adj_hdrs, null, NewChar('A',$x,$y).NewChar(strval($lastrow+2),$x,$y),true); // hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('E',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldCenter); // format hdrs

                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('A',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('A',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldLeft); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('B',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('B',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldCenter); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('C',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('C',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldLeft); // format hdrs
                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($lastrow+2),$x,$y).':'.NewChar('E',$x,$y).NewChar(strval($lastrow+2)),$x,$y)->applyFromArray($BoldCenter); // format hdrs

                    $objPHPExcel->getActiveSheet()->fromArray($data3, null, NewChar('A',$x,$y).NewChar(strval($lastrow+3),$x,$y),true); // 3rd data

                    $firstrow = $lastrow+3;
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                    for ($char='A'; $char <= 'E'; $char++) { 
                        switch ($char) {
                            case ($char == 'B') || ($char == 'D') || ($char == 'E') :
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($CenterSize12);
                                break;
                            default:
                                $objPHPExcel->getActiveSheet()->getStyle(NewChar($char,$x,$y).NewChar(strval(2+$hdr_offset),$x,$y).':'.NewChar($char,$x,$y).NewChar(strval($lastrow),$x,$y))->applyFromArray($Size12);
                                break;
                        }
                    }

                    $objPHPExcel->getActiveSheet()->getStyle(NewChar('D',$x,$y).NewChar(strval($firstrow),$x,$y).':'.NewChar('D',$x,$y).NewChar(strval($lastrow),$x,$y))->getNumberFormat()->setFormatCode('mm/dd/yyyy;;');
                }
                else {
                    $lastrow = $objPHPExcel->getActiveSheet()->getHighestRow(NewChar('A',$x,$y));
                }
                $dataColCount = PHPExcel_Cell::columnIndexFromString($objPHPExcel->getActiveSheet()->getHighestColumn()); // # of data cols
                echo "\r\nAuto Sizing Columns ".NumToAlpha(1)."-".NumToAlpha($dataColCount);
                for ($i = 1; $i <= $dataColCount; $i++) {
                    $objPHPExcel->getActiveSheet()->getColumnDimension(NumToAlpha($i))->setAutoSize(true);
                } 
                $objPHPExcel->GetActiveSheet()->setSelectedCell('A1'); 
            }    
            $Cmds = 'BadCensusDOBs_Template';
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