<?php
    function KandGImpFmtTemplate($template = "None") {
        switch ($template) {
        case "Unknown_BenesysDOBReturn":
            $Cmds = "Cmds_Unknown_BenesysDOBReturn";
            function Cmds_Unknown_BenesysDOBReturn($in_filename) {
                $objReader = PHPExcel_IOFactory::createReader('Excel2007');
                $objPHPExcel = $objReader->load($in_filename);

                $hdrrange = "A1:Q2";
                $ar_hdrrange = explode(":", $hdrrange);
                $startcol = preg_replace('/[0-9]+/',"",$ar_hdrrange[0]);
                $endcol = preg_replace('/[0-9]+/',"",$ar_hdrrange[1]);
                $numcols = AlphaToNum($endcol) - AlphaToNum($startcol) + 1;
                $hdrstrtrow = preg_replace("/[^0-9]/","",$ar_hdrrange[0]);
                $datastrt = preg_replace("/[^0-9]/","",$ar_hdrrange[0]) + count($ar_hdrrange);  
                $AllHdrs = array();

                for ($i=0; $i < count($ar_hdrrange); $i++) { // read header data
                    // echo "Range : ".$startcol.strval($hdrstrtrow+$i).':'.$endcol.strval($hdrstrtrow+$i);
                    $GetHdr = $objPHPExcel->getActiveSheet()->rangeToArray($startcol.strval($hdrstrtrow+$i).':'.$endcol.strval($hdrstrtrow+$i));
                    $AHdr = array();
                    for ($a=0; $a < count($GetHdr[0]); $a++) {
                        array_push($AHdr, $GetHdr[0][$a]);
                    }
                    // var_dump($AHdr);
                    array_push($AllHdrs, $AHdr); 
                }

                $OneHdr = $AllHdrs[0];

                if (count($AllHdrs) > 1) {
                    for ($i=0; $i < (count($AllHdrs)); $i++) { // format header data into 1 array
                        for ($a=0; $a < count($AllHdrs[0]); $a++) { 
                            if ($i<(count($AllHdrs))) {
                                if ($i>0) {
                                    $OneHdr[$a] = $OneHdr[$a].'|r|n'. $AllHdrs[$i][$a];
                                } 
                            }                           
                        } 
                    }
                }

                for ($i=0; $i < 17; $i++) { 
                    $HdrAr1[$i] = $HdrAr1[$i].'|r|n'. $HdrAr2[$i];
                }

                $datastrt = 3;
                $dataend = $objPHPExcel->getActiveSheet()->getHighestRow('A');
                $firstrow = False;
                $lastrow = False;

                 while ($firstrow = True) { // determine first data row
                    $count = 0;
                    $row = $objPHPExcel->getActiveSheet()->getRowIterator($datastrt)->current();
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);
                
                    foreach ($cellIterator as $cell) {
                        $count+=1;
                    }
                    if ($count>$numcols*0.2) { // if row seems to be a data row
                        break;
                    } else {
                        $datastrt+=1;
                    }                    
                }

                while ($lastrow = True) { // determine last data row
                    $count = 0;
                    $row = $objPHPExcel->getActiveSheet()->getRowIterator($dataend)->current();
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);
                
                    foreach ($cellIterator as $cell) {
                        $count+=1;
                    }
                    if ($count>7) { // if row seems to be a data row w/ >7 data cells
                        break;
                    } else {
                        $dataend-=1;
                    }                    
                }
               
                $DataAr = array();
                $DataTypes = array();

                for ($i=AlphaToNum($startcol); $i < ($numcols + 1); $i++) { 
                    $DataColAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NumToAlpha($i).$datastrt.':'.NumToAlpha($i).$dataend));
                    $type = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$datastrt)->getDataType();
                    if ($type == 'null') { // check next cell type in col until type is not null
                        $nextcell = $datastrt+1;
                        while ($type == 'null') {
                            $type = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$nextcell)->getDataType();
                            if ($nextcell >= $dataend) { // give up
                                break;
                            }
                            $nextcell+=1;
                        }
                    }

                    if ($type == 'n' ) { // check if excel float date format, convert
                        if ($i < 10) {
                            $cell = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$datastrt);
                            $InvDate = $cell->getValue();
                            if (PHPExcel_Shared_Date::isDateTime($cell)) { // needs more work...
                                $DataColAr = array_map('TimeStampToDateStr', $DataColAr); // convert Excel Serial Number Format->UNIX timestamp->datestring
                                $type = "Date";
                            }
                        }
                    }

                    array_push($DataAr, $DataColAr);
                    array_push($DataTypes, $type);
                }

                // const TYPE_STRING2  = 'str';
                // const TYPE_STRING   = 's';
                // const TYPE_FORMULA  = 'f';
                // const TYPE_NUMERIC  = 'n';
                // const TYPE_BOOL     = 'b';
                // const TYPE_NULL     = 'null';
                // const TYPE_INLINE   = 'inlineStr'; // Rich text
                // const TYPE_ERROR    = 'e';

                // write Data Store
                $datastorepath = dirname((__FILE__)) .'Datastore'.strval(rand(1000,2000)).'.txt';
                WriteDataStore($datastorepath, $DataAr);
                unset($objPHPExcel);
                unset($objReader);
                // return data into array, hdr arrays, data types...
                return array($datastorepath, $OneHdr, $DataTypes);
            }


            return array($Cmds,True);
            break;
        
        default: // empty arrays (no headers, any formatting, no 'Cmds')
            // echo "Running Default Import ...\r\n";
            $Cmds = "Cmds_Default";
            function Cmds_Default($in_filename , $hdrrange) {
                $objReader = PHPExcel_IOFactory::createReader('Excel2007');
                $objPHPExcel = $objReader->load($in_filename);
                // determine hdrs, data
                $ar_hdrrange = explode(":", $hdrrange);
                $startcol = preg_replace('/[0-9]+/',"",$ar_hdrrange[0]);
                $endcol = preg_replace('/[0-9]+/',"",$ar_hdrrange[1]);
                $numcols = AlphaToNum($endcol) - AlphaToNum($startcol);
                $hdrstrtrow = preg_replace("/[^0-9]/","",$ar_hdrrange[0]);
                $datastrt = preg_replace("/[^0-9]/","",$ar_hdrrange[0]) + count($ar_hdrrange);              

                $AllHdrs = array();
                
                for ($i=0; $i < count($ar_hdrrange); $i++) { // read header data
                    // echo "Range : ".$startcol.strval($hdrstrtrow+$i).':'.$endcol.strval($hdrstrtrow+$i);
                    $GetHdr = $objPHPExcel->getActiveSheet()->rangeToArray($startcol.strval($hdrstrtrow+$i).':'.$endcol.strval($hdrstrtrow+$i));
                    $AHdr = array();
                    for ($a=0; $a < count($GetHdr[0]); $a++) {
                        array_push($AHdr, $GetHdr[0][$a]);
                    }
                    // var_dump($AHdr);
                    array_push($AllHdrs, $AHdr); 
                }

                $OneHdr = $AllHdrs[0];

                if (count($AllHdrs) > 1) {
                    for ($i=0; $i < (count($AllHdrs)); $i++) { // format header data into 1 array
                        for ($a=0; $a < count($AllHdrs[0]); $a++) { 
                            if ($i<(count($AllHdrs))) {
                                if ($i>0) {
                                    $OneHdr[$a] = $OneHdr[$a].'|r|n'. $AllHdrs[$i][$a];
                                } 
                            }                           
                        } 
                    }
                }
             
                $dataend = max($objPHPExcel->getActiveSheet()->getHighestRow('A'), $objPHPExcel->getActiveSheet()->getHighestRow('B'), $objPHPExcel->getActiveSheet()->getHighestRow('C'));

                $firstrow = False;
                $lastrow = False;

                while ($firstrow = True) { // determine first data row
                    $count = 0;
                    $row = $objPHPExcel->getActiveSheet()->getRowIterator($datastrt)->current();
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);
                
                    foreach ($cellIterator as $cell) {
                        $count+=1;
                    }
                    if ($count>$numcols*0.2) { // if row seems to be a data row
                        break;
                    } else {
                        $datastrt+=1;
                    }                    
                }


                while ($lastrow = True) { // determine last data row
                    $count = 0;
                    $row = $objPHPExcel->getActiveSheet()->getRowIterator($dataend)->current();
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);
                
                    foreach ($cellIterator as $cell) {
                        $count+=1;
                    }
                    if ($count>($numcols*0.7)) { // if row seems to be a data row
                        break;
                    } else {
                        $dataend-=1;
                    }                    
                }
               
                $DataAr = array();
                $DataTypes = array();
                for ($i=AlphaToNum($startcol); $i < ($numcols + 1); $i++) { 
                    $DataColAr = CompactRangeArray($objPHPExcel->getActiveSheet()->rangeToArray(NumToAlpha($i).$datastrt.':'.NumToAlpha($i).$dataend));
                    $type = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$datastrt)->getDataType();
                    if ($type == 'null') { // check next cell type in col until type is not null
                        $nextcell = $datastrt+1;
                        while ($type == 'null') {
                            $type = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$nextcell)->getDataType();
                            if ($nextcell >= $dataend) { // give up
                                break;
                            }
                            $nextcell+=1;
                        }
                    }

                    if ($type == 'n' ) { // check if excel float date format, convert
                        $cell = $objPHPExcel->getActiveSheet()->getCell(NumToAlpha($i).$datastrt);
                        $InvDate = $cell->getValue();
                        if (PHPExcel_Shared_Date::isDateTime($cell)) { // needs more work...
                            $DataColAr = array_map('TimeStampToDateStr', $DataColAr); // convert Excel Serial Number Format->UNIX timestamp->datestring
                            $type = "Date";
                        }
                    }

                    array_push($DataAr, $DataColAr);
                    array_push($DataTypes, $type);
                }

                // write Data Store
                $datastorepath = dirname((__FILE__)) .'/Datastore'.strval(rand(1000,2000)).'.txt';
                WriteDataStore($datastorepath, $DataAr);
                unset($objPHPExcel);
                unset($objReader);
                // return data into array, hdr arrays, data types...
                return array($datastorepath, $OneHdr, $DataTypes);
            }
            return array($Cmds,True);
            break;
        }
    }
?>