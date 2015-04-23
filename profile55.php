<? 

include 'modbus_crc.php';

/* Переменные для соединения с базой данных */ 
$hostname = "localhost"; 
$username = "root"; 
$password = "Rfnfgekmrf48"; 
$dbName = "electro"; 

/* Таблица MySQL, в которой хранятся данные */ 
$userstable = "profile55"; 

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


	
// ------------------------------55 Счетчик ----------------------------------

$interval = 16;
$adr_req = "55";
$cod_req = "0603";
$num_req = "0F";

$j = 1;
$k_m = 31;


$time = strtotime("-1 day");
$lastday = date("d.m.Y", $time);
$time1 = '00:00';
$time2 = '23:30';

$period1 = strtotime($lastday)+strtotime($time1);
$period2 = strtotime($lastday)+strtotime($time2);

//echo $period1;
//echo '</br>';
//echo $period2;
//echo '</br>';

$num = ($period2-$period1)/1800;

//echo $num;
//echo '</br>';

/* создать соединение */ 
mysql_connect($hostname,$username,$password) OR DIE("Не могу создать соединение "); 
/* выбрать базу данных. Если произойдет ошибка - вывести ее */ 
mysql_select_db($dbName) or die(mysql_error());  
//echo 'Соединение с базой данных установлено';
//echo '</br>';

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


$timest = strtotime($date_prof)+strtotime($time_prof);
/* составить запрос для вставки информации о данных в таблицу */ 
$query = "INSERT INTO $userstable VALUES(0, $timest,($reactive_prof/2), ($active_prof/2))"; 
/* Выполнить запрос. Если произойдет ошибка - вывести ее. */ 
mysql_query($query) or die(mysql_error()); 
//echo "Информация занесена в базу данных."; 


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

$timest = strtotime($date_prof)+strtotime($time_prof);
/* составить запрос для вставки информации о данных в таблицу */ 
$query = "INSERT INTO $userstable VALUES(0, $timest,($reactive_prof/2), ($active_prof/2))"; 
/* Выполнить запрос. Если произойдет ошибка - вывести ее. */ 
mysql_query($query) or die(mysql_error()); 
//echo "Информация занесена в базу данных."; 

//echo "Дата: ".$date_prof." | Время: ".$time_prof." | Активная мощность: ".($active_prof/2)." | Реактивная мощность: ".($reactive_prof/2);
//echo '</br>';
//echo $j;
//echo '</br>';
$active_array[$j] = $active_prof/2;
$j++;
};
}

dio_close($fd);

// ---------------------------------------------------------------------

/* Закрыть соединение */ 
mysql_close(); 
?> 