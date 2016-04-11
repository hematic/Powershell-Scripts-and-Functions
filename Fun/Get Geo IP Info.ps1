$restapiuri = "http://geoip.nekudo.com/api/173.65.16.57"
$JSONResponse = Invoke-RestMethod -Uri $restapiuri -ContentType application/json -method Get