<?php
    function Wrap_ExptoXLS() { // ** CAN'T PASS variable number of arguments from 4D
        require_once dirname(__FILE__) . '/KG_Exp4DtoXLS.php';

        $arguments = func_get_args(); // list all args instead
        $pathname = $arguments[0];
        $x_offset = $arguments[1];
        $y_offset = $arguments[2];
        $template = $arguments[3];
        $ar_headers = $arguments[4];
        array_splice($arguments, 0, 5);

        $ar_data = array();
        foreach ($arguments as $array) {
            array_push($ar_data, $array);
        }
        
        // *** ADD check pathname function
        // set_error_handler("HandleError");       
        $ar_data = ShiftDataArrays ($ar_data); // preformat data arrays to be 'rows'

        Exp4DtoXls($pathname,$x_offset,$y_offset,$template,$ar_headers,$ar_data);

        return True;
    }

    function HandleError($errno, $errstr) {
        echo "<b>Error:</b> [$errno] $errstr<br>";
        return "<b>Error:</b> [$errno] $errstr<br>";
        exit();
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
?>