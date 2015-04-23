<?php
//$f = fopen('/dev/ttyr01',"r+");
$fd = dio_open('/dev/ttyr01', O_RDWR);

dio_fcntl($fd, F_SETFL, O_SYNC);

dio_tcsetattr($fd, array(
    'baud' => 9600,
    'bits' => 8,
    'stop' => 1,
    'parity' =>0
));

dio_write($fd, "\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81", 11);



usleep(1000);

//dio_write($fd, "x00x05x00x00x10x25",6);

//usleep(200);

//$result = dio_read($fd, 25);

//dio_write($fd, "\x00\x02\x80\x71", 4);
//usleep(200);
//dio_write($fd,"x00x00x01xB0");
$result = dio_read($fd, 15);

//echo $result;

dio_write($fd, "\x00\x05\x50\x00\x2C\x25", 6);

usleep(1000);

$result = dio_read($fd, 64);
$realres = dechex(ord($result)-40);

echo $result;
echo $realres;
dio_close($fd);

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