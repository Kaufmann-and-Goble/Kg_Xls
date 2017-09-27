<?php
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
            function Default_Template($objPHPExcel, $x = 1, $y = 1) {
            }
            $Cmds = 'Default_Template';
            $StyleArrs = array($HdrAr,$HdrStyleAr,$Cmds);
            return $StyleArrs;
            break;
        }
    }
?>