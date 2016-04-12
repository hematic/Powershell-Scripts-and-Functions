$restapiuri = "http://geoip.nekudo.com/api/$Ip"
$JSONResponse = Invoke-RestMethod -Uri $restapiuri -ContentType application/json -method Get