$s="THeQuickBrownFoxJumpsOverTheLazyDog"
$i=24
MsgBox(64,"",StringMid($s,1,$i)&@CRLF&StringMid($s,$i+1,StringLen($s)))
