<!doctype html>

<html lang="ru">
<head>
     <meta charset="utf-8" />
     <title>Газ</title>
	 
<?php


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


	
// ------------------------------Газовый счетчик ----------------------------------

$fd = dio_open('/dev/ttyr03', O_RDWR);
dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 19200,
    'bits' => 7,
    'stop' => 1,
    'parity' => 0
));

dio_write($fd, "x2Fx3Fx21x0Dx0A",5);

//dio_write($fd,chr(10),1);
sleep (10);
echo ('10');

$result = dio_read($fd);

dio_close($fd);
echo $result;
?>

<body>
</body>
</html>