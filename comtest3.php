<?


function usleep_win($msec) {
   usleep($msec * 1000);
}

#exec('mode com3: baud=9600 data=8 stop=1 parity=n xon=off');

$fd = dio_open('/dev/ttyr01', O_RDWR);

###### WIN32 WORKAROUND


##### FOR LINUX BOX

//dio_fcntl($fd, F_SETFL, O_SYNC);


dio_tcsetattr($fd, array(
  'baud' => 9600,
  'bits' => 8,
  'stop'  => 1,
  'parity' => 0
)); 


##############

function merc_gd($cmd, $factor = 1, $total = 0)
{
	global $fd;
	global $sleep_time;

	usleep_win(50);
	flush();
	dio_write($fd, $cmd, 6);
	usleep_win($sleep_time);
	$result = dio_read($fd, 64);

	$ret = array();
	
	if ( $total != 1 )
	{
		if ( dechex(ord($result[7])) >= 40 )
		$result[7] = chr(dechex(ord($result[7])) - 40);
		if ( dechex(ord($result[10])) >= 40 )
		$result[10] = chr(dechex(ord($result[10])) - 40);
		if ( dechex(ord($result[13])) >= 40 )
		$result[13] = chr(dechex(ord($result[13])) - 40);
		if ( dechex(ord($result[16])) >= 40 )
		$result[16] = chr(dechex(ord($result[16])) - 40);

		$ret[0] = hexdec(dd($result[7]).dd($result[9]).dd($result[8]))*$factor;
		if ( strlen($result) > 12 )
		$ret[1] = hexdec(dd($result[10]).dd($result[12]).dd($result[11]))*$factor;
		if ( strlen($result) > 15 )
		$ret[2] = hexdec(dd($result[13]).dd($result[15]).dd($result[14]))*$factor;
		if ( strlen($result) > 18 )
		$ret[3] = hexdec(dd($result[16]).dd($result[18]).dd($result[17]))*$factor;
	}
	else
	$ret[0] = hexdec(dd($result[8]).dd($result[7]).dd($result[10]).dd($result[9]))*$factor;


	return $ret;
}





$sleep_time = 200;

function dd($data = "")
{
	$result = "";
	$data2 = "";
	for ( $j = 0; $j < count($data); $j++ )
	{
	echo $j;
		$data2 = dechex(ord($data[0]));
		if ( strlen($data2) == 1  )
		$result = "0".$data2;
		else
		$result .= $data2;

	}
	return $result;
}


# Инициализация соединения, передача пароля
dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);
usleep_win($sleep_time);
$result = dio_read($fd, 15);

//for ( $i = 0; $i < 25; $i++ )
$n = 0;
$total_cnt = 90;
$total_cnt1 = 1080;
while ( $n == 0 )
	{
echo $n;
	# Сила тока по фазам
	# =====================================================
	$Ia = merc_gd("\x00\x08\x16\x21\x4F\x9E", 0.001);
	echo "Ia: $Ia[0] - $Ia[1] - $Ia[2]\r\n";
	$It = $Ia[0] + $Ia[1] + $Ia[2];
/*	# Мощность по фазам
	# =====================================================
	$Pv = merc_gd("\x00\x08\x16\x00\x8F\x86", 0.01);
	if ( round($Pv[0], 2) != round($Pv[1] + $Pv[2] + $Pv[3], 2) )
	$error = "error"; else $error = "";
	echo "Pv: $Pv[0] - $Pv[1] - $Pv[2] - $Pv[3] $error\r\n";
	# Cosf по фазам
	# =====================================================
	$Cos = merc_gd("\x00\x08\x16\x30\x8F\x92", 0.001);
	echo "Cos: $Cos[0] - $Cos[1] - $Cos[2] - $Cos[3]\r\n";
	# Напряжение по фазам
	# =====================================================
	$Uv = merc_gd("\x00\x08\x16\x11\x4F\x8A", 0.01);
	echo "Uv: $Uv[0] - $Uv[1] - $Uv[2]\r\n";

	if ( empty($Uv[0]) && empty($Uv[1]) && empty($Uv[2]) )
	{
		dio_close($fd);
		$fd = dio_open('/dev/ttyr01', O_RDWR);
		dio_tcsetattr($fd, array(
		  'baud' => 9600,
		  'bits' => 8,
		  'stop'  => 1,
		  'parity' => 0
		)); 

		sleep(2);

		# Инициализация соединения, передача пароля
		dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);
		usleep_win($sleep_time);
		$result = dio_read($fd, 15);

	}

	$total_cnt++;
	if ( $total_cnt >= 90 )
	{
		$Tot = merc_gd("\x00\x05\x00\x00\x10\x25", 0.001, 1);
		echo "Total: $Tot[0]\r\n";

		// За текущие сутки
		$Tot = merc_gd("\x00\x05\x40\x00\x21\xE5", 0.001, 1);
		echo "Total cur: $Tot[0]\r\n";

		$total_cnt = 0;
	}

	$total_cnt1++;
	if ( $total_cnt1 >= 1080 )
	{
		// За предыдущие сутки
		$Tot = merc_gd("\x00\x05\x50\x00\x2C\x25", 0.001, 1);
		echo "Total prev: $Tot[0]\r\n";

		$total_cnt1 = 0;
	}

*/
	sleep(10);
}

# Завершение соединения
dio_write($fd, "\x00\x02\x80\x71", 4);
usleep_win($sleep_time);
$result = dio_read($fd, 8);
dio_close($fd);

?>
