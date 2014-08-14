<?php
header('Content-Type:text/html;charset=UTF-8');
libxml_use_internal_errors(true);
$doc = new DOMDocument();
$doc->loadHtmlFile('http://www.esunbank.com.tw/event/announce/BankCode.htm');
$nowRow = array();
$nextRow = array();
$index = 0;
echo "\"銀行代號\", \"銀行名稱\", \"類型\"\n";
foreach ($doc->getElementsByTagName('td') as $e) {
	$index ++;
	$nowRow[$index] = $e->nodeValue;
	if($index <= 24) continue;
	if ($index % 2 == 0) {
		if ($nowRow[$index] != '')
			echo "\"".$nowRow[$index - 1]."\", \"".$nowRow[$index]."\", \"".getBandType($nowRow[$index - 1])."\"\n";
	}
	
}
function getBandType($type) {
	if ($type == 108 || $type == 118 || $type == 147 || $type == 151) {
		return '銀行';
	}
	if ($type == 104 || $type == 106 || $type == 114 || $type == 115) {
		return '信用合作社';
	}
	if ( $type < 104 || ( $type >= 803 && $type <= 822)) {
		return '銀行';
	}
	if ( $type < 104 || ( $type >= 803 && $type <= 822)) {
		return '銀行';
	}
	if ( $type >= 104 && $type <= 224) {
		return '銀行';
	}
	return '農會';
}
?>