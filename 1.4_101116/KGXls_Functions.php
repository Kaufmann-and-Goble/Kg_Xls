<?php
    function Wrap_ExptoXLS($filepath, $x_offset = 1, $y_offset = 1, $template = "", $ar_headers = array(), $ar_types = array(), $Datapath ) {
        require_once dirname(__FILE__) . '/KG_Exp4DtoXLS.php';
        // $filepath = str_replace(' ', '\ ', $filepath);

        if (!is_Dir(dirname($filepath))) { // if immediate directory doesnt exist, create it
            if (!mkdir(dirname($filepath), 0777, true)) { // if directory could not be created
                exit('Could not Create Directory(s) for : '.$filepath);
                }
        }

        if (!file_exists($Datapath)) {
            exit('DataFile Does Not Exist : '.$Datapath);
        }

        $ar_data = ReadDataStore($Datapath, $ar_types);

        // $ar_data = ShiftDataArrays ($ar_data); // preformat data arrays to be 'rows'
        // set_error_handler("HandleError");       

        Exp4DtoXls($filepath,$x_offset,$y_offset,$template,$ar_headers,$ar_data);
        unlink($Datapath); // delete data store
        return True;
    }

    function Wrap_Impto4D($filepath, $template = "", $hdrrange = "") {
        require_once dirname(__FILE__) . '/KG_ImpXLSto4D.php';
        // $filepath = str_replace(' ', '\ ', $filepath);

        if (!is_Dir(dirname($filepath))) { // if immediate directory doesnt exist, create it
                exit('Could Access filepath : '.$filepath);
        }
        $ar_data = ImpXlsto4D($filepath,$template,$hdrrange);
        echo "^^".$ar_data[0]."~~";
        echo json_encode($ar_data[1]) ."~~";
        echo json_encode($ar_data[2]) ."~~";
    }

    function ReadDataStore ($Datapath, $types = array()) { // read and return data arrays from csv formatted file
        $handle = fopen($Datapath, "r"); // read only
        $data_ars = array();
        if ($handle) {
            $idx = 0;
            while (($buffer = fgets($handle)) !== false) { 
                $array = array();
                $array = str_replace("~~", ",", str_getcsv ($buffer, ",", "", "|")); // b/c Delimiter "|" not working
                switch($types{$idx}) { // convert to appropriate data type
                    case 16 :
                        $array = array_map('intval', $array);
                        break; // longint
                    case 6:
                        $array = array_map('boolval', $array);
                        break; // bool
                    case 14:
                        $array = array_map('floatval', $array);
                        break; // real
                    case 17:
                        $array = array_map('strtotime', $array);
                        break; // date
                    default:
                        $array = array_map('strval', $array);
                        break; // text
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

    function WriteDataStore ($Datafilepath, $Data) { // write data arrays from XLS file to text file
        $handle = fopen($Datafilepath, 'w') or die('Cannot open file:  '.$Datafilepath);
        $numcols = count($Data);
        $countcols = 1;
        foreach($Data as $col) {
            for ($i=0; $i < count($col); $i++) {
                fwrite($handle, $col[$i]."~|~");
            }
            if ($countcols<$numcols) {
                fwrite($handle, "\n");
            }            
            $countcols+=1;
        }
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

    function CheckDateFormat ($string) { // looking for mm/dd/yyyy
        $dt = DateTime::createFromFormat('m/d/Y', $string);
        return $dt !== false && !array_sum($dt->getLastErrors());
    }

    function DateStrToTimeStamp ($string) { // must be mm/dd/yyyy
        $datetime = DateTime::createFromFormat('m/d/Y', $string);
        $xlsDate = PHPExcel_Shared_Date::PHPToExcel($datetime);
        return floor($xlsDate); // remove time
    }

    function TimeStampToDateStr($float) { //excel float -> mm/dd/yyyy string
        $datestr = PHPExcel_Style_NumberFormat::toFormattedString($float, 'MM/DD/YYYY');
        // $datestr = PHPExcel_Style_NumberFormat::toFormattedString($float, 'YYYY/DD/MM');
        return ($datestr);
    }

    function ExplicitStr ($string) {
        return strval($string);
    }

    function MultiHdrAdjust ($hdrs = array()) { // create 'multi-line' headers if applicable
        // $Headers{1}:="Social|r|nSecurity" $Headers{2}:="Last|r|nName"
        // -> $Headers{1}:=Social $Headers{2}:=Last , $Headers2{1}:=Security $Headers2{2}:=Name
        $hdrs_out = array();
        $MultiRow = False;
        $numlines = 0;

        for ($i=0; $i < count($hdrs); $i++) {
            if (substr_count($hdrs[$i], '|r|n') > $numlines) {
                $numlines = substr_count($hdrs[$i], '|r|n');
            }
        }
        $numlines = $numlines+1;
        array_push($hdrs_out,$numlines-1);

        $myarray = array();
        for ($line=0; $line < $numlines; $line++) { 
            $newarray = array();
            array_push($myarray,$newarray);
        }

        for ($i=0; $i < count($hdrs); $i++) {
            $pos = 0;
            $strtpos = 0;
            $NoMore = False;
            $len = strlen($hdrs[$i]);
            // echo "header: ".$hdrs[$i]."\r\n";
            // echo "length: ".$len."\r\n";
            for ($line=0; $line < $numlines; $line++) {
                $pos = strpos($hdrs[$i],'|r|n', $strtpos);
                // echo "pos: ".$pos."\r\n";
                // echo "strtpos: ".$strtpos."\r\n";
                    if ($pos === false)  {
                        if ($strtpos == 0 & $NoMore != True) {
                            // echo "Whole Header"."\r\n";
                            array_push($myarray[$line],$hdrs[$i]);
                            $NoMore = True;
                        }
                        else if ($strtpos > 0 & $strtpos+3 < $len & $NoMore != True) {
                            array_push($myarray[$line],substr($hdrs[$i], $strtpos+3));
                            // echo "Substring2: ".substr($hdrs[$i], $strtpos+3)."\r\n";
                            $NoMore = True;
                        }
                        else {
                            // echo "Nothing"."\r\n";
                            array_push($myarray[$line],'');
                        }
                    }
                    else {
                        if ($strtpos>0) {
                            $strtpos+=3;
                        }
                        array_push($myarray[$line],substr($hdrs[$i], $strtpos, $pos - $strtpos));
                        // echo "Substring: ".substr($hdrs[$i], $strtpos, $pos - $strtpos)."\r\n";
                        $pos+=1;
                        $strtpos=$pos;
                    }
            }
            
        }
        for ($line=0; $line < $numlines; $line++) { 
            array_push($hdrs_out,$myarray[$line]);
        }
        return $hdrs_out;
    }

    function MultiHdrAdjust2 ($hdrs = array()) { // create mult sets of headers if applicable
        // $Headers{1}:="Social|r|nSecurity" $Headers{2}:="Last|r|nName"
        // -> $Headers{1}:=Social $Headers{2}:=Last , $Headers2{1}:=Security $Headers2{2}:=Name
        $hdrs1 = array();
        $hdrs2 = array();
        $hdrs_out = array();
        $MultiRow = False;
        for ($i=0; $i < count($hdrs); $i++) {
            $pos = strpos($hdrs[$i],'|r|n');
            if ($pos === false) {
                array_push($hdrs1,$hdrs[$i]);
                array_push($hdrs2,'');
            }
            else {
                array_push($hdrs1,substr($hdrs[$i], 0, $pos));
                array_push($hdrs2,substr($hdrs[$i], $pos + 4));
                $MultiRow = True;
            }             
        }
        if ($MultiRow) {
            array_push($hdrs_out,1,$hdrs1,$hdrs2);
        }
        else {
            array_push($hdrs_out,0,$hdrs1);
        }
        var_dump($hdrs_out);
        return $hdrs_out;
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