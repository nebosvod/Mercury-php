<?
$fd = dio_open('/dev/ttyr01', O_RDWR);

dio_fcntl($fd, F_SETFL, O_SYNC);


dio_tcsetattr($fd, array(
  'baud' => 9600,
  'bits' => 8,
  'stop'  => 1,
  'parity' => 0
));


function merc_gd($cmd, $factor = 1, $total = 0)
{
	global $fd;
	global $sleep_time;

	usleep(50);
	flush();
	dio_write($fd, $cmd, 6);
	usleep($sleep_time);
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
		$data2 = dechex(ord($data[0]));
		if ( strlen($data2) == 1  )
		$result = "0".$data2;
		else
		$result .= $data2;

	}
	return $result;
}


dio_write($fd, "x00x01x01x01x01x01x01x01x01x77x81", 11);
usleep($sleep_time);
$result = dio_read($fd, 15);

$n = 0;
$total_cnt = 90;
$total_cnt1 = 370;
while ( $n == 0 )
{

	
	$Ia = merc_gd("x00x08x16x21x4Fx9E", 0.001);
	echo "Ia: $Ia[0] - $Ia[1] - $Ia[2]";
	# Мощность по фазам
	# =====================================================
	$Pv = merc_gd("x00x08x16x00x8Fx86", 0.01);
	echo "Pv: $Pv[0] - $Pv[1] - $Pv[2] - $Pv[3] $error";
	# Cosf по фазам
	# =====================================================
	$Cos = merc_gd("x00x08x16x30x8Fx92", 0.001);
	echo "Cos: $Cos[0] - $Cos[1] - $Cos[2] - $Cos[3]";
	# Напряжение по фазам
	# =====================================================
	$Uv = merc_gd("x00x08x16x11x4Fx8A", 0.01);
	echo "Uv: $Uv[0] - $Uv[1] - $Uv[2]";

	$total_cnt++;
	# Каждые 15 минут собираем информацию о текущем потреблении за сутки
	if ( $total_cnt >= 90 )
	{
		$Tot = merc_gd("x00x05x00x00x10x25", 0.001, 1);
		echo "Total: $Tot[0]";

		// За текущие сутки
		$Tot = merc_gd("x00x05x40x00x21xE5", 0.001, 1);
		echo "Total cur: $Tot[0]";
		$total_cnt = 0;
	}

	$total_cnt1++;
	if ( $total_cnt1 >= 370 )
	{
		// За предыдущие сутки
		$Tot = merc_gd("x00x05x50x00x2Cx25", 0.001, 1);
		echo "Total prev: $Tot[0]";

		$total_cnt1 = 0;
	}


	sleep(10);
}

# Завершение соединения
dio_write($fd, "x00x02x80x71", 4);
usleep($sleep_time);
$result = dio_read($fd, 8);
dio_close($fd);

?>