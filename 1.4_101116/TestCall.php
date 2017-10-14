<?php
	require_once dirname((__FILE__)) . '/KG_ImpXLSto4D.php';
	$path = strval(dirname(dirname((__FILE__)))) . '/DropFolder/393_UnknownAddress_20170901_BeneSys.xlsx';
	// /Volumes/ME_DCPs/Import_Export/For_Testing/db_CoOp/DevDropFolder/393_UnknownAddress_20170901_BeneSys.xlsx
	KG_ImpXLSto4D($path, "Unknown_BenesysDOBReturn")

?>