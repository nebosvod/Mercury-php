<html lang="ru">
<head>
     <meta charset="utf-8" />
     <title>Профиль мощности</title>

<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');
/** Include PHPExcel */
require_once 'Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();


// Add some data
$objPHPExcel->setActiveSheetIndex(0);
$aSheet = $objPHPExcel->getActiveSheet();
$aSheet->getColumnDimension('A')->setWidth(10);
$aSheet->getColumnDimension('B')->setWidth(25);
$aSheet->getColumnDimension('C')->setWidth(15);
$aSheet->getColumnDimension('D')->setWidth(15);
$aSheet->getColumnDimension('E')->setWidth(15);
$aSheet->getColumnDimension('F')->setWidth(15);
$aSheet->getColumnDimension('G')->setWidth(15);

$aSheet->mergeCells('A1:G1');
$aSheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('A1')->getFont()->setBold(true);
$aSheet->setCellValue('A1','СВЕДЕНИЯ');

$aSheet->mergeCells('A2:G2');
$aSheet->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('A2')->getFont()->setBold(true);
$aSheet->setCellValue('A2','о расходе электроэнергии по ОАО "Аньковское", за Февраль 2013 года.');

$aSheet->mergeCells('A3:A4');
$aSheet->mergeCells('B3:B4');
$aSheet->mergeCells('C3:C4');
$aSheet->mergeCells('D3:E3');
$aSheet->mergeCells('F3:F4');
$aSheet->mergeCells('G3:G4');
$aSheet->getStyle('B3:G3')->getAlignment()->setWrapText(true);
$aSheet->getRowDimension('3')->setRowHeight(30);

$aSheet->getStyle('A3:G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('A3:G3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->getStyle('A3:G3')->getFont()->setBold(true);
$aSheet->setCellValue('A3','№ п\п');
$aSheet->setCellValue('B3','Наименование точки учета');
$aSheet->setCellValue('C3','Номер прибора учета');
$aSheet->setCellValue('D3','Показания прибора учета');
$aSheet->setCellValue('F3','Расчетный к-т');
$aSheet->setCellValue('G3','Расход э/энергии (кВтч)');

$aSheet->getStyle('D4:E4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('D4:E4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->getStyle('D4:E4')->getFont()->setBold(true);
$aSheet->setCellValue('D4','Конечное');
$aSheet->setCellValue('E4','Начальное');

$aSheet->getStyle('A5:G5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('A5:G5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$aSheet->setCellValue('A5','1');
$aSheet->setCellValue('B5','2');
$aSheet->setCellValue('C5','3');
$aSheet->setCellValue('D5','4');
$aSheet->setCellValue('E5','5');
$aSheet->setCellValue('F5','6');
$aSheet->setCellValue('G5','7');

$aSheet->getStyle('A6:G27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('A6:G27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->getStyle('A5:G27')->getFont()->setBold(true);
$aSheet->setCellValue('A6','1. ');
$aSheet->mergeCells('B6:G6');
$aSheet->getStyle('B6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B6','Потребление электроэнергии Потребителем');

$aSheet->getRowDimension('7')->setRowHeight(30);
$aSheet->getRowDimension('8')->setRowHeight(30);
$aSheet->getRowDimension('9')->setRowHeight(30);
$aSheet->getRowDimension('10')->setRowHeight(30);
$aSheet->getRowDimension('11')->setRowHeight(30);
$aSheet->getRowDimension('15')->setRowHeight(30);

$aSheet->getStyle('B7:B11')->getAlignment()->setWrapText(true);

$aSheet->setCellValue('A7','1.1 ');
$aSheet->setCellValue('A8','1.2 ');
$aSheet->setCellValue('A9','1.3 ');
$aSheet->setCellValue('A10','1.4 ');
$aSheet->setCellValue('A11','1.5 ');

$aSheet->setCellValue('B7','"Аньково" ВЛ-102 ЗТП-400 Т-1 активный');
$aSheet->setCellValue('B8','"Аньково" ВЛ-102 ЗТП-400 Т-1 реактивный');
$aSheet->setCellValue('B9','"Аньково" ВЛ-102 ЗТП-400 Т-2 активный');
$aSheet->setCellValue('B10','"Аньково" ВЛ-102 ЗТП-400 Т-2 реактивный');
$aSheet->setCellValue('B11','"Аньково" ВЛ-102 КТП-400 Очистные сооружения');

$aSheet->setCellValue('C7','00335385 ');
$aSheet->setCellValue('C8','00335385 ');
$aSheet->setCellValue('C9','01753886-07 ');
$aSheet->setCellValue('C10','01753886-07 ');
$aSheet->setCellValue('C11','0447906609 ');

$aSheet->setCellValue('F7','120 ');
$aSheet->setCellValue('F8','120');
$aSheet->setCellValue('F9','120');
$aSheet->setCellValue('F10','120');
$aSheet->setCellValue('F11','20');

$aSheet->setCellValue('A12','2. ');
$aSheet->mergeCells('B12:F12');
$aSheet->getStyle('B12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B12','Потерии э/э оплачиваемые активные');
$aSheet->setCellValue('A13','2.1 ');

$aSheet->setCellValue('A14','3. ');
$aSheet->mergeCells('B14:F14');
$aSheet->getStyle('B14')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B14','Потерии э/э оплачиваемые реактивные');
$aSheet->setCellValue('A15','3.1 ');
$aSheet->getStyle('B15')->getAlignment()->setWrapText(true);
$aSheet->setCellValue('B15','Реактивное потребление по "Правилу"');
$aSheet->setCellValue('A16','3.2 ');
$aSheet->setCellValue('A17','4. ');
$aSheet->mergeCells('B17:G17');
$aSheet->getStyle('B17')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B17','Замена приборов учета э/э');
$aSheet->setCellValue('A18','4.1 ');
$aSheet->setCellValue('A19','4.2 ');
$aSheet->mergeCells('B20:F20');
$aSheet->getStyle('B20')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B20','ИТОГО ОБЩИЙ РАСХОД (С ПОТЕРЯМИ), АКТИВНЫЙ');
$aSheet->mergeCells('B21:F21');
$aSheet->getStyle('B21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('B21','ИТОГО ОБЩИЙ РАСХОД (С ПОТЕРЯМИ), РЕАКТИВНЫЙ');
$aSheet->getStyle('A23:G27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$aSheet->setCellValue('F23','М. П.');
$aSheet->setCellValue('B25','Исполнитель__________________Смолярчук А.Н.');
$aSheet->setCellValue('B27','телефон (4932) 33102');



//$styleArray = array( 'borders' => array( 'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$styleArray2 = array( 'borders' => array( 'allborders' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$aSheet->getStyle('A3:G21')->applyFromArray($styleArray2);


include 'modbus_crc.php';

// Преобразование ascii в hex---------------
    function ascii2hex($ascii)
    {
    $hex = '';
	for ($i = 0; $i < strlen($ascii); $i++)
	{
	$byte = strtoupper(dechex(ord($ascii{$i})));
	$byte = str_repeat('0', 2 - strlen($byte)).$byte;
	$hex.=$byte;
	};
    return $hex;
    };
//-------------------------------------------

    function ans_explode ($prof)
    {
	global $active_mounth, $reactive_mounth;
	
	
	$active_mounth = substr($prof,4,2).substr($prof,2,2).substr($prof,8,2).substr($prof,6,2);
	$reactive_mounth = substr($prof,20,2).substr($prof,18,2).substr($prof,24,2).substr($prof,22,2);
	$active_mounth = hexdec($active_mounth)/1000;
	$reactive_mounth = hexdec($reactive_mounth)/1000;
    };





$mounth = '12';

if ($mounth == '01')
	{
		$req_begin = '060202AA10';
		$req_mounth = '053100';
	};

if ($mounth == '02')
	{
		$req_begin = '060202FF10';
		$req_mounth = '053200';
	};	

if ($mounth == '03')
	{
		$req_begin = '0602035410';
		$req_mounth = '053300';
	};

if ($mounth == '04')
	{
		$req_begin = '060203A910';
		$req_mounth = '053400';
	};

if ($mounth == '05')
	{
		$req_begin = '060203FE10';
		$req_mounth = '053500';
	};

if ($mounth == '06')
	{
		$req_begin = '0602045310';
		$req_mounth = '053600';
	};

if ($mounth == '07')
	{
		$req_begin = '060204A810';
		$req_mounth = '053700';
	};	
	
if ($mounth == '08')
	{
		$req_begin = '060204FD10';
		$req_mounth = '053800';
	};	

if ($mounth == '09')
	{
		$req_begin = '0602055210';
		$req_mounth = '053900';
	};	

if ($mounth == '10')
	{
		$req_begin = '060205A710';
		$req_mounth = '053A00';
	};	

if ($mounth == '11')
	{
		$req_begin = '060205FC10';
		$req_mounth = '053B00';
	};		

if ($mounth == '12')
	{
		$req_begin = '0602065110';
		$req_mounth = '053C00';
	};


//--------------Обработка 55 счетчика

$adr_req = '55';

$fd = dio_open('/dev/ttyr01', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x55\x01\x01\x01\x01\x01\x01\x01\x01\xB4\xD2", 11);
$result = dio_read($fd, 15);


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';

$aSheet->setCellValue('E7',round(($active_mounth/2),3));
$aSheet->setCellValue('E8',round(($reactive_mounth/2),3));

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo (ascii2hex($result2));
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';

$aSheet->setCellValue('G7',($active_mounth*120));
$aSheet->setCellValue('G8',($reactive_mounth*120));

dio_close($fd);



//--------------Обработка 56 счетчика

$adr_req = '56';

$fd = dio_open('/dev/ttyr00', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x56\x01\x01\x01\x01\x01\x01\x01\x01\xA0\x22", 11);
$result = dio_read($fd, 15);


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';


$aSheet->setCellValue('E9',round(($active_mounth/2),3));
$aSheet->setCellValue('E10',round(($reactive_mounth/2),3));


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo (ascii2hex($result2));
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';

$aSheet->setCellValue('G9',($active_mounth*120));
$aSheet->setCellValue('G10',($reactive_mounth*120));

dio_close($fd);

//--------------Обработка 42 счетчика

$adr_req = '42';

$fd = dio_open('/dev/ttyr02', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x42\x01\x01\x01\x01\x01\x01\x01\x01\x5F\x22", 11);
$result = dio_read($fd, 15);


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';

$aSheet->setCellValue('E11',round(($active_mounth/2),3));


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
echo $req;
echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
echo $crc_req1;
echo '</br>';
echo $crc_req2;
echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode (ascii2hex($result2));
echo (ascii2hex($result2));
ans_explode (ascii2hex($result2));
echo round($active_mounth,3);
echo '</br>';
echo round($reactive_mounth,3);
echo '</br>';

$aSheet->setCellValue('G11',($active_mounth*20));

dio_close($fd);

$aSheet->setCellValue('D7', '=G7/120+E7');
$aSheet->setCellValue('D8', '=G8/120+E8');
$aSheet->setCellValue('D9', '=G9/120+E9');
$aSheet->setCellValue('D10', '=G10/120+E10');
$aSheet->setCellValue('D11', '=G11/20+E11');
$aSheet->setCellValue('G12', '3777');
$aSheet->setCellValue('G14', '20464');
$aSheet->setCellValue('G20', '=G7+G9+G11+G12');
$aSheet->setCellValue('G21', '=G8+G10+G14');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('/home/public/'.$mounth_otch.'-'.$year_otch.'-расход.xls');
$objWriter->save('/home/public/расход.xls');
?>
<body>
</body>
</html>