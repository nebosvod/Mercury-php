<!doctype html>

<html lang="en">
<head>
     <meta charset="utf-8" />
     <title>jQuery UI Datepicker - Select a Date Range</title>
     <script src="/jquery/js/jquery-1.8.2.js"></script>
     <script src="/jquery/js/jquery-ui-1.9.1.custom.js"></script>
     <link rel="stylesheet" href="/jquery/css/redmond/jquery-ui-1.9.1.custom.css" />
     <script src="/jquery/development-bundle/ui/i18n/jquery.ui.datepicker-ru.js"></script>
     <script>
	 $(function() {
	 $("#datepicker").datepicker($.datepicker.regional["ru"]);
         $( "#from" ).datepicker({
         defaultDate: "+1w",
         changeMonth: true,
         onClose: function( selectedDate ) {
         $( "#to" ).datepicker( "option", "minDate", selectedDate );
    		 }
         });
         $( "#to" ).datepicker({
         defaultDate: "+1w",
         changeMonth: true,
         onClose: function( selectedDate ) {
         $( "#from" ).datepicker( "option", "maxDate", selectedDate );
         }
         });
     });
    
     </script>
 </head>
 <body>
    <form name = 'dateform' action = 'profile.php' method = 'post'>
 <label for="from">From</label>
 <input type="text" id="from" name="from" />
 <label for="to">to</label>
 <input type="text" id="to" name="to" />
 <label for="temepick">Время</label>
 <input type="text" id="teimepick" name="timepick" />
 <br />
 <input type = 'submit' value = 'Отправить' />
    </form>
</body>
</html>

<?php

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
	$time = substr($ans,8,4);
	$date = substr($ans,12,2).".".substr($ans,14,2).".".substr($ans,16,2);
	$answer = substr($ans,0,20);
	$crc = substr($ans,20,4);
    };
// ----------------------------------------------------------------




// Вычисление разницы во времени --------------------------------------

    function time_dif($timestampl, $timestamp2)
    {
    if ($timestampl && $timestamp2)
	{
	$time_lapse = $timestamp2 - $timestampl;
	return $time_lapse;
	};
    return false; 
    };
// ---------------------------------------------------------------------------

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
echo ascii2hex($result);

echo "---------";
dio_write($fd, "\x55\x08\x13\x27\xDD", 5);

usleep(1000);

$result1 = dio_read($fd, 64);

echo ascii2hex($result1);

ans_explode(ascii2hex($result1));
echo "<br>";
echo ("Сетевой адрес счетчика - ".$adress);
echo "<br>";
echo ("Текущий адрес памяти - ".$memory);
echo "<br>";
echo ("Дата - ".$date);
echo "<br>";
echo ("Время - ".$time);
echo "<br>";
echo ("CRC - ".$crc);
echo "<br>";
echo ("Ответ - ".$answer);

crc16(chr(0x55));
echo "<br>";
echo $res;
dio_close($fd);

?>