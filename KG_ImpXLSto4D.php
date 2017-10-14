<?php
        require_once dirname((__FILE__)) . '/Classes/PHPExcel.php';
        require_once dirname((__FILE__)) . '/KG_ImpFormats.php'; // Retrieve Format Function

        function ImpXlsto4D($in_filename, $template = "None", $hdrrange = "") {
            // parmtype check
            set_time_limit(600);

            $templateArrays = KandGImpFmtTemplate($template); //retrieve formats
            $TemplateCmds = $templateArrays[0]; // retrieve Commands
            $StandardExp = $templateArrays[1]; // Standard Import
            // Template-Specific 'Commands'
            // echo "\r\nExecuting Template Commands: \r\nCmd: ".$TemplateCmds."\r\n";
            $data = call_user_func($TemplateCmds, $in_filename, $hdrrange);  // Needs an Error Handler
            return $data;
        }
?>