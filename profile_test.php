<!doctype html>

<html lang="ru">
<head>
     <meta charset="utf-8" />
     <title>Профиль мощности</title>
	 
<?php

/*
/** Error reporting */
//error_reporting(E_ALL);
//ini_set('display_errors', TRUE);
//ini_set('display_startup_errors', TRUE);
//date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');
/** Include PHPExcel */
require_once 'Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();


// Add some data
$objPHPExcel->setActiveSheetIndex(0);
$aSheet = $objPHPExcel->getActiveSheet();
$aSheet->getColumnDimension('A')->setWidth(24);
$aSheet->getColumnDimension('B')->setWidth(12);
$aSheet->getColumnDimension('C')->setWidth(12);
$aSheet->getColumnDimension('D')->setWidth(12);
$aSheet->getColumnDimension('E')->setWidth(12);
$aSheet->getColumnDimension('F')->setWidth(12);
$aSheet->getColumnDimension('G')->setWidth(12);
$aSheet->getColumnDimension('H')->setWidth(12);
$aSheet->getColumnDimension('I')->setWidth(12);
$aSheet->getColumnDimension('J')->setWidth(12);
$aSheet->getColumnDimension('K')->setWidth(12);
$aSheet->getColumnDimension('L')->setWidth(12);
$aSheet->getColumnDimension('M')->setWidth(12);
$aSheet->getColumnDimension('N')->setWidth(12);
$aSheet->getColumnDimension('O')->setWidth(12);
$aSheet->getColumnDimension('P')->setWidth(12);
$aSheet->getColumnDimension('Q')->setWidth(12);
$aSheet->getColumnDimension('R')->setWidth(12);
$aSheet->getColumnDimension('S')->setWidth(12);
$aSheet->getColumnDimension('T')->setWidth(12);
$aSheet->getColumnDimension('U')->setWidth(12);
$aSheet->getColumnDimension('V')->setWidth(12);
$aSheet->getColumnDimension('W')->setWidth(12);
$aSheet->getColumnDimension('X')->setWidth(12);
$aSheet->getColumnDimension('Y')->setWidth(12);
$aSheet->getColumnDimension('Z')->setWidth(12);

$aSheet->mergeCells('B5:B6');
$aSheet->mergeCells('C5:Z5');

$aSheet->getStyle('A1')->getFont()->setBold(true);
$aSheet->setCellValue('A1','ОАО "Аньковское"');
//$aSheet->getStyle('A2')->getFont()->setBold(true);
//$aSheet->setCellValue('A2','Отчет за потребленную электроэнергию и мощность, ноябрь 2012г');
$aSheet->setCellValue('A3','Счетчик №');
$aSheet->getStyle('B3')->getFont()->setBold(true);
$aSheet->setCellValue('B3','335385');
$aSheet->setCellValue('A4','Тр.тока (коэф)');
$aSheet->getStyle('B4')->getFont()->setBold(true);
$aSheet->setCellValue('B4','120');
$aSheet->getStyle('C5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->setCellValue('C5','Время суток');
$aSheet->getRowDimension('6')->setRowHeight(30);
$aSheet->getStyle('B5:B37')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('B41:B71')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('B75:B105')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C6:Z6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C6:Z6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->getStyle('B5')->getAlignment()->setWrapText(true);
$aSheet->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
$aSheet->setCellValue('B5','Число расчетного месяца');
$aSheet->setCellValue('C6','0.00-1.00');
$aSheet->setCellValue('D6','1.00-2.00');
$aSheet->setCellValue('E6','2.00-3.00');
$aSheet->setCellValue('F6','3.00-4.00');
$aSheet->setCellValue('G6','4.00-5.00');
$aSheet->setCellValue('H6','5.00-6.00');
$aSheet->setCellValue('I6','6.00-7.00');
$aSheet->setCellValue('J6','7.00-8.00');
$aSheet->setCellValue('K6','8.00-9.00');
$aSheet->setCellValue('L6','9.00-10.00');
$aSheet->setCellValue('M6','10.00-11.00');
$aSheet->setCellValue('N6','11.00-12.00');
$aSheet->setCellValue('O6','12.00-13.00');
$aSheet->setCellValue('P6','13.00-14.00');
$aSheet->setCellValue('Q6','14.00-15.00');
$aSheet->setCellValue('R6','15.00-16.00');
$aSheet->setCellValue('S6','16.00-17.00');
$aSheet->setCellValue('T6','17.00-18.00');
$aSheet->setCellValue('U6','18.00-19.00');
$aSheet->setCellValue('V6','19.00-20.00');
$aSheet->setCellValue('W6','20.00-21.00');
$aSheet->setCellValue('X6','21.00-22.00');
$aSheet->setCellValue('Y6','22.00-23.00');
$aSheet->setCellValue('Z6','23.00-24.00');
//$aSheet->getColumnDimension('AA5')->setWidth(18);
//$aSheet->getStyle('AA5')->getFont()->setBold(true);
//$aSheet->setCellValue('AA5','Макс. знач.');

$aSheet->getStyle('Y38')->getFont()->setBold(true);
$aSheet->setCellValue('Y38','ИТОГО');

$aSheet->setCellValue('A39','Счетчик №');
$aSheet->getStyle('B39')->getFont()->setBold(true);
$aSheet->setCellValue('B39','1753886');
$aSheet->setCellValue('A40','Тр.тока (коэф)');
$aSheet->getStyle('B40')->getFont()->setBold(true);
$aSheet->setCellValue('B40','120');
$aSheet->getStyle('Y72')->getFont()->setBold(true);
$aSheet->setCellValue('Y72','ИТОГО');


$aSheet->setCellValue('B7','1');
$aSheet->setCellValue('B8','2');
$aSheet->setCellValue('B9','3');
$aSheet->setCellValue('B10','4');
$aSheet->setCellValue('B11','5');
$aSheet->setCellValue('B12','6');
$aSheet->setCellValue('B13','7');
$aSheet->setCellValue('B14','8');
$aSheet->setCellValue('B15','9');
$aSheet->setCellValue('B16','10');
$aSheet->setCellValue('B17','11');
$aSheet->setCellValue('B18','12');
$aSheet->setCellValue('B19','13');
$aSheet->setCellValue('B20','14');
$aSheet->setCellValue('B21','15');
$aSheet->setCellValue('B22','16');
$aSheet->setCellValue('B23','17');
$aSheet->setCellValue('B24','18');
$aSheet->setCellValue('B25','19');
$aSheet->setCellValue('B26','20');
$aSheet->setCellValue('B27','21');
$aSheet->setCellValue('B28','22');
$aSheet->setCellValue('B29','23');
$aSheet->setCellValue('B30','24');
$aSheet->setCellValue('B31','25');
$aSheet->setCellValue('B32','26');
$aSheet->setCellValue('B33','27');
$aSheet->setCellValue('B34','28');
$aSheet->setCellValue('B35','29');
$aSheet->setCellValue('B36','30');
$aSheet->setCellValue('B37','31');

$aSheet->setCellValue('B41','1');
$aSheet->setCellValue('B42','2');
$aSheet->setCellValue('B43','3');
$aSheet->setCellValue('B44','4');
$aSheet->setCellValue('B45','5');
$aSheet->setCellValue('B46','6');
$aSheet->setCellValue('B47','7');
$aSheet->setCellValue('B48','8');
$aSheet->setCellValue('B49','9');
$aSheet->setCellValue('B50','10');
$aSheet->setCellValue('B51','11');
$aSheet->setCellValue('B52','12');
$aSheet->setCellValue('B53','13');
$aSheet->setCellValue('B54','14');
$aSheet->setCellValue('B55','15');
$aSheet->setCellValue('B56','16');
$aSheet->setCellValue('B57','17');
$aSheet->setCellValue('B58','18');
$aSheet->setCellValue('B59','19');
$aSheet->setCellValue('B60','20');
$aSheet->setCellValue('B61','21');
$aSheet->setCellValue('B62','22');
$aSheet->setCellValue('B63','23');
$aSheet->setCellValue('B64','24');
$aSheet->setCellValue('B65','25');
$aSheet->setCellValue('B66','26');
$aSheet->setCellValue('B67','27');
$aSheet->setCellValue('B68','28');
$aSheet->setCellValue('B69','29');
$aSheet->setCellValue('B70','30');
$aSheet->setCellValue('B71','31');

$aSheet->setCellValue('A73','Счетчик №');
$aSheet->getStyle('B73')->getFont()->setBold(true);
$aSheet->setCellValue('B73','4479066');
$aSheet->setCellValue('A74','Тр.тока (коэф)');
$aSheet->getStyle('B74')->getFont()->setBold(true);
$aSheet->setCellValue('B74','20');
$aSheet->getStyle('Y106')->getFont()->setBold(true);
$aSheet->setCellValue('Y106','ИТОГО');



$aSheet->setCellValue('B75','1');
$aSheet->setCellValue('B76','2');
$aSheet->setCellValue('B77','3');
$aSheet->setCellValue('B78','4');
$aSheet->setCellValue('B79','5');
$aSheet->setCellValue('B80','6');
$aSheet->setCellValue('B81','7');
$aSheet->setCellValue('B82','8');
$aSheet->setCellValue('B83','9');
$aSheet->setCellValue('B84','10');
$aSheet->setCellValue('B85','11');
$aSheet->setCellValue('B86','12');
$aSheet->setCellValue('B87','13');
$aSheet->setCellValue('B88','14');
$aSheet->setCellValue('B89','15');
$aSheet->setCellValue('B90','16');
$aSheet->setCellValue('B91','17');
$aSheet->setCellValue('B92','18');
$aSheet->setCellValue('B93','19');
$aSheet->setCellValue('B94','20');
$aSheet->setCellValue('B95','21');
$aSheet->setCellValue('B96','22');
$aSheet->setCellValue('B97','23');
$aSheet->setCellValue('B98','24');
$aSheet->setCellValue('B99','25');
$aSheet->setCellValue('B100','26');
$aSheet->setCellValue('B101','27');
$aSheet->setCellValue('B102','28');
$aSheet->setCellValue('B103','29');
$aSheet->setCellValue('B104','30');
$aSheet->setCellValue('B105','31');

$aSheet->setCellValue('B109','И.о. главного энергетика__________________Смолярчук А.Н.');

$styleArray = array( 'borders' => array( 'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$styleArray2 = array( 'borders' => array( 'allborders' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$aSheet->getStyle('C7:Z37')->applyFromArray($styleArray2);
$aSheet->getStyle('C41:Z71')->applyFromArray($styleArray2);
$aSheet->getStyle('C75:Z105')->applyFromArray($styleArray2);

$aSheet->getStyle('B6:Z6')->applyFromArray($styleArray2);
$aSheet->getStyle('B5:Z5')->applyFromArray($styleArray2);
$aSheet->getStyle('B7:B37')->applyFromArray($styleArray2);
$aSheet->getStyle('B41:B71')->applyFromArray($styleArray2);
$aSheet->getStyle('B75:B105')->applyFromArray($styleArray2);

$aSheet->getStyle('Y38:Z38')->applyFromArray($styleArray2);
$aSheet->getStyle('Y72:Z72')->applyFromArray($styleArray2);
$aSheet->getStyle('Y106:Z106')->applyFromArray($styleArray2);

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



// Разбивка ответа текущего состояния ячейки памяти хранения профиля мощности, времени и даты -----

    function ans_explode ($ans)
    {
	global $adress, $memory, $time, $date, $answer, $crc;
	$adress = substr($ans,0,2);
	$memory = substr($ans,2,4);
	$time = substr($ans,8,2).":".substr($ans,10,2);
	$date = substr($ans,12,2).".".substr($ans,14,2).".20".substr($ans,16,2);
	$answer = substr($ans,0,20);
	$crc = substr($ans,20,4);
    };
// ----------------------------------------------------------------

    function ans_explode2 ($prof2)
    {
	global $active_mounth, $reactive_mounth;
	
	
	$active_mounth = substr($prof2,4,2).substr($prof2,2,2).substr($prof2,8,2).substr($prof2,6,2);
	$reactive_mounth = substr($prof2,20,2).substr($prof2,18,2).substr($prof2,24,2).substr($prof2,22,2);
	$active_mounth = hexdec($active_mounth)/1000;
	$reactive_mounth = hexdec($reactive_mounth)/1000;
    };

// Разбивка ответа показаний профилей мощностей -----

    function ans_explode_prof ($prof)
    {
	global $adress_prof, $time_prof, $date_prof, $period_prof, $active_prof, $reactive_prof, $crc_prof;
	
	$adress_prof = substr($prof,0,2);
	$time_prof = substr($prof,4,2).":".substr($prof,6,2);
	$date_prof = substr($prof,8,2).".".substr($prof,10,2).".20".substr($prof,12,2);
	$active_prof = substr($prof,18,2).substr($prof,16,2);
	$reactive_prof = substr($prof,26,2).substr($prof,24,2);
	$crc_prof = substr($prof,32,4);
	$active_prof = hexdec($active_prof)/1000;
	$reactive_prof = hexdec($reactive_prof)/1000;
    };


	
 //------------------------------55 Счетчик ----------------------------------
echo '--------------Профиль мощности счетчика №335385----------------------------';
echo '</br>';


$interval = 16;
$adr_req = "55";
$cod_req = "0603";
$num_req = "0F";

$j = 1;
$k_m = 31;

$date1 = $_POST['date1'];
$date2 = $_POST['date2'];
$time1 = $_POST['time1'];
$time2 = $_POST['time2'];

$mounth = substr($date1,3,2);
$year_otch = substr($date1,6,4);

if ($mounth == '01') $mounth_otch = 'Январь'; 
if ($mounth == '02') $mounth_otch = 'Февраль';
if ($mounth == '03') $mounth_otch = 'Март';
if ($mounth == '04') $mounth_otch = 'Апрель';
if ($mounth == '05') $mounth_otch = 'Май';
if ($mounth == '06') $mounth_otch = 'Июнь';
if ($mounth == '07') $mounth_otch = 'Июль';
if ($mounth == '08') $mounth_otch = 'Август';
if ($mounth == '09') $mounth_otch = 'Сентябрь';
if ($mounth == '10') $mounth_otch = 'Октябрь';
if ($mounth == '11') $mounth_otch = 'Ноябрь';
if ($mounth == '12') $mounth_otch = 'Декабрь';

if (($mounth == '01') or ($mounth == '03') or ($mounth == '05') or ($mounth == '07') or ($mounth == '08') or ($mounth == '10') or ($mounth == '12')) $k_m = 31;
if (($mounth == '04') or ($mounth == '06') or ($mounth == '09') or ($mounth == '11')) $k_m = 30;
if (($mounth == '02') and ($year_otch == '2013')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2014')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2015')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2016')) $k_m = 29;
if (($mounth == '02') and ($year_otch == '2017')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2018')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2019')) $k_m = 28;
if (($mounth == '02') and ($year_otch == '2020')) $k_m = 29;

$aSheet->getStyle('A2')->getFont()->setBold(true);
$txt_otch = 'Отчет за потребленную электроэнергию и мощность, '.$mounth_otch.' '.$year_otch.' г.';
$aSheet->setCellValue('A2','Отчет за потребленную электроэнергию и мощность, '.$mounth_otch.' '.$year_otch.' г.');

$period1 = strtotime($date1)+strtotime($time1);
$period2 = strtotime($date2)+strtotime($time2);

$num = ($period2-$period1)/1800;

$fd = dio_open('/dev/ttyr01', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);
$result = dio_read($fd, 15);

dio_write($fd, "\x55\x08\x13\x27\xDD", 5);


$result1 = dio_read($fd, 64);


ans_explode(ascii2hex($result1));

$tek = strtotime($date)+strtotime($time);

$num_tek = ($tek-$period1)/1800;

crc16(chr(0x55));

for ($i=($num_tek-1); $i>=($num_tek-$num-1); $i--)
{
if ((hexdec($memory)-$interval*$i) <= 0)
{
$mem_req =strtoupper(dechex((65520+(hexdec($memory)-$interval*$i))));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
echo $mem_req;
echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));

echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
echo '</br>';
echo $j;
echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};


if ((hexdec($memory)-$interval*$i) >= 0)
{
$mem_req = strtoupper(dechex(hexdec($memory)-$interval*$i));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
echo $mem_req;
echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));
echo (ascii2hex($result2));
echo '</br>';
echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
echo '</br>';
echo $j;
echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};
}

dio_close($fd);

$k1 = 1;

for ($k=1; $k<=$k_m; $k++)
{
$aSheet->setCellValue('C'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('D'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('E'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('F'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('G'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('H'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('I'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('J'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('K'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('L'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('M'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('N'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('O'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('P'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Q'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('R'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('S'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('T'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('U'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('V'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('W'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('X'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Y'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Z'.($k+6),($active_array[$k1++]+$active_array[$k1++]));
};

$aSheet->getStyle('Z38')->getFont()->setBold(true);
$aSheet->setCellValue('Z38', '=SUM(C7:Z37)*B4');



//--------------------------------56 счетчик--------------------------------------
//echo '--------------Профиль мощности счетчика №1753886----------------------------';
//echo '</br>';

$interval = 16;
$adr_req = "56";
$cod_req = "0603";
$num_req = "0F";

$j = 1;

$date1 = $_POST['date1'];
$date2 = $_POST['date2'];
$time1 = $_POST['time1'];
$time2 = $_POST['time2'];

$period1 = strtotime($date1)+strtotime($time1);
$period2 = strtotime($date2)+strtotime($time2);

$num = ($period2-$period1)/1800;


$fd = dio_open('/dev/ttyr00', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);
$result = dio_read($fd, 15);

dio_write($fd, "\x56\x08\x13\xD7\xDD", 5);


$result1 = dio_read($fd, 64);


ans_explode(ascii2hex($result1));

$tek = strtotime($date)+strtotime($time);

$num_tek = ($tek-$period1)/1800;

crc16(chr(0x56));

for ($i=($num_tek-1); $i>=($num_tek-$num-1); $i--)
{
if ((hexdec($memory)-$interval*$i) <= 0)
{
$mem_req =strtoupper(dechex((65520+(hexdec($memory)-$interval*$i))));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
//echo $mem_req;
//echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));

//echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
//echo '</br>';
//echo $j;
//echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};


if ((hexdec($memory)-$interval*$i) >= 0)
{
$mem_req = strtoupper(dechex(hexdec($memory)-$interval*$i));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
//echo $mem_req;
//echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));
//echo (ascii2hex($result2));
//echo '</br>';
//echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
//echo '</br>';
//echo $j;
//echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};
}

dio_close($fd);

$k1 = 1;

for ($k=1; $k<=$k_m; $k++)
{
$aSheet->setCellValue('C'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('D'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('E'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('F'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('G'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('H'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('I'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('J'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('K'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('L'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('M'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('N'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('O'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('P'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Q'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('R'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('S'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('T'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('U'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('V'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('W'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('X'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Y'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Z'.($k+40),($active_array[$k1++]+$active_array[$k1++]));
};

for ($k=1; $k<=1440; $k++)
{
$sum_active56 = $sum_active56 + $active_array[$k++];
}

$aSheet->getStyle('Z72')->getFont()->setBold(true);
$aSheet->setCellValue('Z72', '=SUM(C41:Z71)*B40');


//--------------------------------42 счетчик--------------------------------------
//echo '--------------Профиль мощности счетчика №4479066----------------------------';
//echo '</br>';

$interval = 16;
$adr_req = "42";
$cod_req = "0603";
$num_req = "0F";

$j = 1;

$date1 = $_POST['date1'];
$date2 = $_POST['date2'];
$time1 = $_POST['time1'];
$time2 = $_POST['time2'];

$period1 = strtotime($date1)+strtotime($time1);
$period2 = strtotime($date2)+strtotime($time2);

$num = ($period2-$period1)/1800;


$fd = dio_open('/dev/ttyr02', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);
$result = dio_read($fd, 15);

dio_write($fd, "\x42\x08\x13\x97\xD9", 5);


$result1 = dio_read($fd, 64);


ans_explode(ascii2hex($result1));

$tek = strtotime($date)+strtotime($time);

$num_tek = ($tek-$period1)/1800;

crc16(chr(0x42));

for ($i=($num_tek-1); $i>=($num_tek-$num-1); $i--)
{
if ((hexdec($memory)-$interval*$i) <= 0)
{
$mem_req =strtoupper(dechex((65520+(hexdec($memory)-$interval*$i))));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
//echo $mem_req;
//echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));
//echo (ascii2hex($result2));
//echo '</br>';
//echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
//echo '</br>';
//echo $j;
//echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};


if ((hexdec($memory)-$interval*$i) >= 0)
{
$mem_req = strtoupper(dechex(hexdec($memory)-$interval*$i));
if (strlen($mem_req) == 1)
{
$mem_req = '000'.$mem_req;
}
if (strlen($mem_req) == 2)
{
$mem_req = '00'.$mem_req;
}
if (strlen($mem_req) == 3)
{
$mem_req = '0'.$mem_req;
}
//echo $mem_req;
//echo '</br>';


$cod_req1 = substr($cod_req,0,2);
$cod_req2 = substr($cod_req,2,4);
$mem_req1 = substr($mem_req,0,2);
$mem_req2 = substr($mem_req,2,4);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req);

$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.$cod_req1).chr('0x'.$cod_req2).chr('0x'.$mem_req1).chr('0x'.$mem_req2).chr('0x'.$num_req).chr('0x'.$crc_req1).chr('0x'.$crc_req2);

dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode_prof (ascii2hex($result2));

//echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".$reactive_prof;
//echo '</br>';
//echo $j;
//echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};
}

dio_close($fd);

$k1 = 1;

for ($k=1; $k<=$k_m; $k++)
{
$aSheet->setCellValue('C'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('D'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('E'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('F'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('G'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('H'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('I'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('J'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('K'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('L'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('M'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('N'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('O'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('P'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Q'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('R'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('S'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('T'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('U'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('V'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('W'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('X'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Y'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
$aSheet->setCellValue('Z'.($k+74),($active_array[$k1++]+$active_array[$k1++]));
};

$aSheet->getStyle('Z106')->getFont()->setBold(true);
$aSheet->setCellValue('Z106', '=SUM(C75:Z105)*B74');

// Create a new worksheet called “My Data”
$myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'Summa');
// Attach the “My Data” worksheet as the first worksheet in the PHPExcel object
$objPHPExcel->addSheet($myWorkSheet, 1);
$objPHPExcel->setActiveSheetIndex(1);
$aSheet = $objPHPExcel->getActiveSheet();

$aSheet->getColumnDimension('A')->setWidth(3);
$aSheet->getColumnDimension('B')->setWidth(12);
$aSheet->getColumnDimension('C')->setWidth(10);
$aSheet->getColumnDimension('D')->setWidth(10);
$aSheet->getColumnDimension('E')->setWidth(10);
$aSheet->getColumnDimension('F')->setWidth(10);
$aSheet->getColumnDimension('G')->setWidth(10);
$aSheet->getColumnDimension('H')->setWidth(10);
$aSheet->getColumnDimension('I')->setWidth(10);
$aSheet->getColumnDimension('J')->setWidth(10);
$aSheet->getColumnDimension('K')->setWidth(10);
$aSheet->getColumnDimension('L')->setWidth(10);
$aSheet->getColumnDimension('M')->setWidth(10);
$aSheet->getColumnDimension('N')->setWidth(10);
$aSheet->getColumnDimension('O')->setWidth(10);
$aSheet->getColumnDimension('P')->setWidth(10);
$aSheet->getColumnDimension('Q')->setWidth(10);
$aSheet->getColumnDimension('R')->setWidth(10);
$aSheet->getColumnDimension('S')->setWidth(10);
$aSheet->getColumnDimension('T')->setWidth(10);
$aSheet->getColumnDimension('U')->setWidth(10);
$aSheet->getColumnDimension('V')->setWidth(10);
$aSheet->getColumnDimension('W')->setWidth(10);
$aSheet->getColumnDimension('X')->setWidth(10);
$aSheet->getColumnDimension('Y')->setWidth(10);
$aSheet->getColumnDimension('Z')->setWidth(10);

$aSheet->mergeCells('B5:B6');
$aSheet->mergeCells('C5:Z5');
$aSheet->mergeCells('A1:Z1');
$aSheet->mergeCells('G2:T2');
$aSheet->mergeCells('B3:D3');
$aSheet->mergeCells('X3:Z3');
//$aSheet->mergeCells('B41:F41');
$aSheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->setCellValue('A1','Сведения');
$aSheet->getStyle('G2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->setCellValue('G2','о фактическом почасовом расходе электрической энергии за '.$mounth_otch.' 2013 года ОАО "Аньковское"');
$aSheet->setCellValue('B3','Договор №29');
$aSheet->setCellValue('X3','Уровень СН2');
//$aSheet->setCellValue('B41','Руководитель предприятия');
$aSheet->setCellValue('Q41','М.П.');
$aSheet->getStyle('C5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->setCellValue('C5','Время суток');
$aSheet->getRowDimension('6')->setRowHeight(30);
$aSheet->getStyle('B5:B37')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('B41:B71')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('B75:B105')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C6:Z6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$aSheet->getStyle('C6:Z6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$aSheet->getStyle('B5')->getAlignment()->setWrapText(true);
$aSheet->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
$aSheet->setCellValue('B5','Число расчетного месяца');
$aSheet->setCellValue('C6','0.00-1.00');
$aSheet->setCellValue('D6','1.00-2.00');
$aSheet->setCellValue('E6','2.00-3.00');
$aSheet->setCellValue('F6','3.00-4.00');
$aSheet->setCellValue('G6','4.00-5.00');
$aSheet->setCellValue('H6','5.00-6.00');
$aSheet->setCellValue('I6','6.00-7.00');
$aSheet->setCellValue('J6','7.00-8.00');
$aSheet->setCellValue('K6','8.00-9.00');
$aSheet->setCellValue('L6','9.00-10.00');
$aSheet->setCellValue('M6','10.00-11.00');
$aSheet->setCellValue('N6','11.00-12.00');
$aSheet->setCellValue('O6','12.00-13.00');
$aSheet->setCellValue('P6','13.00-14.00');
$aSheet->setCellValue('Q6','14.00-15.00');
$aSheet->setCellValue('R6','15.00-16.00');
$aSheet->setCellValue('S6','16.00-17.00');
$aSheet->setCellValue('T6','17.00-18.00');
$aSheet->setCellValue('U6','18.00-19.00');
$aSheet->setCellValue('V6','19.00-20.00');
$aSheet->setCellValue('W6','20.00-21.00');
$aSheet->setCellValue('X6','21.00-22.00');
$aSheet->setCellValue('Y6','22.00-23.00');
$aSheet->setCellValue('Z6','23.00-24.00');

$aSheet->setCellValue('B7','1');
$aSheet->setCellValue('B8','2');
$aSheet->setCellValue('B9','3');
$aSheet->setCellValue('B10','4');
$aSheet->setCellValue('B11','5');
$aSheet->setCellValue('B12','6');
$aSheet->setCellValue('B13','7');
$aSheet->setCellValue('B14','8');
$aSheet->setCellValue('B15','9');
$aSheet->setCellValue('B16','10');
$aSheet->setCellValue('B17','11');
$aSheet->setCellValue('B18','12');
$aSheet->setCellValue('B19','13');
$aSheet->setCellValue('B20','14');
$aSheet->setCellValue('B21','15');
$aSheet->setCellValue('B22','16');
$aSheet->setCellValue('B23','17');
$aSheet->setCellValue('B24','18');
$aSheet->setCellValue('B25','19');
$aSheet->setCellValue('B26','20');
$aSheet->setCellValue('B27','21');
$aSheet->setCellValue('B28','22');
$aSheet->setCellValue('B29','23');
$aSheet->setCellValue('B30','24');
$aSheet->setCellValue('B31','25');
$aSheet->setCellValue('B32','26');
$aSheet->setCellValue('B33','27');
$aSheet->setCellValue('B34','28');
$aSheet->setCellValue('B35','29');
$aSheet->setCellValue('B36','30');
$aSheet->setCellValue('B37','31');

$aSheet->setCellValue('B41','И.о. главного энергетика__________________Смолярчук А.Н.');


for ($k=1; $k<=31; $k++)
{
$aSheet->setCellValue('C'.($k+6),('=Worksheet!C'.($k+6).'*120+Worksheet!C'.($k+40).'*120+Worksheet!C'.($k+74).'*20'));
$aSheet->setCellValue('D'.($k+6),('=Worksheet!D'.($k+6).'*120+Worksheet!D'.($k+40).'*120+Worksheet!D'.($k+74).'*20'));
$aSheet->setCellValue('E'.($k+6),('=Worksheet!E'.($k+6).'*120+Worksheet!E'.($k+40).'*120+Worksheet!E'.($k+74).'*20'));
$aSheet->setCellValue('F'.($k+6),('=Worksheet!F'.($k+6).'*120+Worksheet!F'.($k+40).'*120+Worksheet!F'.($k+74).'*20'));
$aSheet->setCellValue('G'.($k+6),('=Worksheet!G'.($k+6).'*120+Worksheet!G'.($k+40).'*120+Worksheet!G'.($k+74).'*20'));
$aSheet->setCellValue('H'.($k+6),('=Worksheet!H'.($k+6).'*120+Worksheet!H'.($k+40).'*120+Worksheet!H'.($k+74).'*20'));
$aSheet->setCellValue('I'.($k+6),('=Worksheet!I'.($k+6).'*120+Worksheet!I'.($k+40).'*120+Worksheet!I'.($k+74).'*20'));
$aSheet->setCellValue('J'.($k+6),('=Worksheet!J'.($k+6).'*120+Worksheet!J'.($k+40).'*120+Worksheet!J'.($k+74).'*20'));
$aSheet->setCellValue('K'.($k+6),('=Worksheet!K'.($k+6).'*120+Worksheet!K'.($k+40).'*120+Worksheet!K'.($k+74).'*20'));
$aSheet->setCellValue('L'.($k+6),('=Worksheet!L'.($k+6).'*120+Worksheet!L'.($k+40).'*120+Worksheet!L'.($k+74).'*20'));
$aSheet->setCellValue('M'.($k+6),('=Worksheet!M'.($k+6).'*120+Worksheet!M'.($k+40).'*120+Worksheet!M'.($k+74).'*20'));
$aSheet->setCellValue('N'.($k+6),('=Worksheet!N'.($k+6).'*120+Worksheet!N'.($k+40).'*120+Worksheet!N'.($k+74).'*20'));
$aSheet->setCellValue('O'.($k+6),('=Worksheet!O'.($k+6).'*120+Worksheet!O'.($k+40).'*120+Worksheet!O'.($k+74).'*20'));
$aSheet->setCellValue('P'.($k+6),('=Worksheet!P'.($k+6).'*120+Worksheet!P'.($k+40).'*120+Worksheet!P'.($k+74).'*20'));
$aSheet->setCellValue('Q'.($k+6),('=Worksheet!Q'.($k+6).'*120+Worksheet!Q'.($k+40).'*120+Worksheet!Q'.($k+74).'*20'));
$aSheet->setCellValue('R'.($k+6),('=Worksheet!R'.($k+6).'*120+Worksheet!R'.($k+40).'*120+Worksheet!R'.($k+74).'*20'));
$aSheet->setCellValue('S'.($k+6),('=Worksheet!S'.($k+6).'*120+Worksheet!S'.($k+40).'*120+Worksheet!S'.($k+74).'*20'));
$aSheet->setCellValue('T'.($k+6),('=Worksheet!T'.($k+6).'*120+Worksheet!T'.($k+40).'*120+Worksheet!T'.($k+74).'*20'));
$aSheet->setCellValue('U'.($k+6),('=Worksheet!U'.($k+6).'*120+Worksheet!U'.($k+40).'*120+Worksheet!U'.($k+74).'*20'));
$aSheet->setCellValue('V'.($k+6),('=Worksheet!V'.($k+6).'*120+Worksheet!V'.($k+40).'*120+Worksheet!V'.($k+74).'*20'));
$aSheet->setCellValue('W'.($k+6),('=Worksheet!W'.($k+6).'*120+Worksheet!W'.($k+40).'*120+Worksheet!W'.($k+74).'*20'));
$aSheet->setCellValue('X'.($k+6),('=Worksheet!X'.($k+6).'*120+Worksheet!X'.($k+40).'*120+Worksheet!X'.($k+74).'*20'));
$aSheet->setCellValue('Y'.($k+6),('=Worksheet!Y'.($k+6).'*120+Worksheet!Y'.($k+40).'*120+Worksheet!Y'.($k+74).'*20'));
$aSheet->setCellValue('Z'.($k+6),('=Worksheet!Z'.($k+6).'*120+Worksheet!Z'.($k+40).'*120+Worksheet!Z'.($k+74).'*20'));
};
$aSheet->getStyle('Y38')->getFont()->setBold(true);
$aSheet->setCellValue('Y38','Сумма');
$aSheet->getStyle('Z38')->getFont()->setBold(true);
$aSheet->setCellValue('Z38','=SUM(C7:Z37)');

//$styleArray = array( 'borders' => array( 'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$styleArray2 = array( 'borders' => array( 'allborders' => array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), ), ), );
$aSheet->getStyle('C7:Z37')->applyFromArray($styleArray2);
$aSheet->getStyle('B6:Z6')->applyFromArray($styleArray2);
$aSheet->getStyle('B5:Z5')->applyFromArray($styleArray2);
$aSheet->getStyle('B7:B37')->applyFromArray($styleArray2);
$aSheet->getStyle('Y38:Z38')->applyFromArray($styleArray2);


$myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'Rashod');
$objPHPExcel->addSheet($myWorkSheet, 2);
$objPHPExcel->setActiveSheetIndex(2);
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
$aSheet->setCellValue('A2','о расходе электроэнергии по ОАО "Аньковское", за '.$mounth_otch.' 2013 года.');

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



$aSheet->setCellValue('D7', '=G7/120+E7');
$aSheet->setCellValue('D8', '=G8/120+E8');
$aSheet->setCellValue('D9', '=G9/120+E9');
$aSheet->setCellValue('D10', '=G10/120+E10');
$aSheet->setCellValue('D11', '=G11/20+E11');
$aSheet->setCellValue('G12', '3777');
$aSheet->setCellValue('G14', '20464');
$aSheet->setCellValue('G20', '=G7+G9+G11+G12');
$aSheet->setCellValue('G21', '=G8+G10+G14');

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
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';

$aSheet->setCellValue('E7',round(($active_mounth/2),3));
$aSheet->setCellValue('E8',round(($reactive_mounth/2),3));

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo (ascii2hex($result2));
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';

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
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';


$aSheet->setCellValue('E9',round(($active_mounth/2),3));
$aSheet->setCellValue('E10',round(($reactive_mounth/2),3));


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo (ascii2hex($result2));
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';

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
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_begin,0,2))).chr('0x'.(substr($req_begin,2,2))).chr('0x'.(substr($req_begin,4,2))).chr('0x'.(substr($req_begin,6,2))).chr('0x'.(substr($req_begin,8,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 8);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';

$aSheet->setCellValue('E11',round(($active_mounth/2),3));


$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2)));
//echo $req;
//echo '</br>';
$crc_req = crc16($req);
$crc_req1 = substr($crc_req,0,2);
$crc_req2 = substr($crc_req,2,2);

$req = chr('0x'.$adr_req).chr('0x'.(substr($req_mounth,0,2))).chr('0x'.(substr($req_mounth,2,2))).chr('0x'.(substr($req_mounth,4,2))).chr('0x'.$crc_req1).chr('0x'.$crc_req2);
//echo $crc_req1;
//echo '</br>';
//echo $crc_req2;
//echo '</br>';
dio_write($fd, $req, 6);
$result2 = dio_read($fd, 64);
ans_explode2 (ascii2hex($result2));
//echo (ascii2hex($result2));
ans_explode2 (ascii2hex($result2));
//echo round($active_mounth,3);
//echo '</br>';
//echo round($reactive_mounth,3);
//echo '</br>';

echo 'Отчет сформирован!';


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


$objPHPExcel->setActiveSheetIndex(0);



$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('/home/public/otch/'.$mounth_otch.'-'.$year_otch.'-Профиль мощности.xls');
//$objWriter->save('/media/res2/111.xls');


exit;

?>
<body>
</body>
</html>