<?php
    function Wrap_ExptoXLS($filepath, $x_offset = 1, $y_offset = 1, $template = "", $ar_headers = array(), $ar_types = array(), $Datapath ) {
        require_once dirname(__FILE__) . '/KG_Exp4DtoXLS.php';

        if (!is_Dir(dirname($filepath))) { // if immediate directory doesnt exist, create it
            if (!mkdir(dirname($filepath), 0777, true)) { // if directory could not be created
                exit('Could not Create Directory(s) for : '.$Datapath);
                }
        }

        if (!file_exists($Datapath)) {
            exit('DataFile Does Not Exist : '.$Datapath);
        }

        $ar_data = ReadDataStore($Datapath, $ar_types);

        // $ar_data = ShiftDataArrays ($ar_data); // preformat data arrays to be 'rows'
        // set_error_handler("HandleError");       

        Exp4DtoXls($filepath,$x_offset,$y_offset,$template,$ar_headers,$ar_data);

        return True;
    }

    function ReadDataStore ($Datapath, $types = array()) { // read and return data arrays from csv formatted file
        $handle = fopen($Datapath, "r"); //read only
        $data_ars = array();
        if ($handle) {
            $idx = 0;
            while (($buffer = fgets($handle)) !== false) { 
                $array = array();
                $array = str_replace("~~", ",", str_getcsv ($buffer, ",", "", "|")); // b/c Delimiter "|" not working
                switch($types{$idx}) { // convert to appropriate data type
                    case 16 :
                        $array = array_map('intval', $array);
                        break; //longint
                    case 6:
                        $array = array_map('boolval', $array);
                        break; //bool
                    case 14:
                        $array = array_map('floatval', $array);
                        break; //real
                    case 17:
                        $array = array_map('strtotime', $array);
                        break; //date
                    default:
                        $array = array_map('strval', $array);
                        break; //text
                }
                array_push($data_ars,$array);
                $idx += 1;
            }
        if (!feof($handle)) {
            echo "Error: Unexpected EOL position\r\n";
        }
        fclose($handle);
        }
        else {
            exit('Could Not Open DataFile: : '.$Datapath);  
        }
        return $data_ars;
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
    function NewString ($string, $x = 0 , $y = 0) { // assumes $string is always formatted properly with $strt and $end chars
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

    function NewChar ($string, $x = 0 , $y = 0) {
        if ($x > 0) {
            $x -= 1; 
        }
        if ($y > 0) {
            $y -= 1;
        }
  
        if (is_numeric($string)) { // numbers
            $string = strval(intval($string)+ $y);
        }
        elseif (!preg_match('/[^A-Za-z]/',$string)) { // letters 
            $string = NumToAlpha(AlphaToNum($string)+$x);
        }
        else
            echo $string." could not be converted.\r\n";
        return $string;            
    }

    function CompactRangeArray ($data = array()) {
        $array1D = array();
        foreach ($data as $i=>$o) { // convert to 1d array
            $array1D[$i] = $o[0];
        }
        return $array1D;
    }

    function DateStrToTimeStamp ($string) { // must be mm/dd/yyyy
        $datetime = DateTime::createFromFormat('m/d/Y' , $string);
        $xlsDate = PHPExcel_Shared_Date::PHPToExcel($datetime);
        return floor($xlsDate); // remove time
    }

    function ShiftDataArrays ($data = array()) { // convert 'col' data arrays to 'row' arrays
        $rows = 0;
        $rowarrays = array();
        foreach($data as $array) { // get largest array
        	if (count($array) > $rows) {
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