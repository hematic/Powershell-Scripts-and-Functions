# First install this module from Prateek Singh
# Details here: https://geekeefy.wordpress.com/2016/05/17/powershell-module-for-google-map/

# From an elevated prompt
Install-Module GoogleMap

# or from a non elevated prompt
Install-Module GoogleMap -Scope CurrentUser

# Before you can fully use the google maps API you need to sign up with google, get an API key, and then enable four features on that API key (geolocation, directions, geocoding, and places)
# Having done that, set these four environment variables equal to that one API key.

$env:GoogleGeoloc_API_Key = "AIzaSyBQM1LDfO3-GP5DDM3NYTISE76NXoRKcHQ" # enabled at https://developers.google.com/maps/documentation/geolocation/get-api-key
$env:GoogleDirection_API_Key = "AIzaSyBQM1LDfO3-GP5DDM3NYTISE76NXoRKcHQ" # enabled at https://developers.google.com/maps/documentation/directions/get-api-key
$env:GoogleGeoCode_API_Key = "AIzaSyBQM1LDfO3-GP5DDM3NYTISE76NXoRKcHQ" # enabled at http://developers.google.com/maps/documentation/geocoding/get-api-key
$env:GooglePlaces_API_Key = "AIzaSyBQM1LDfO3-GP5DDM3NYTISE76NXoRKcHQ" # Enabled at https://developers.google.com/places/web-service/get-api-key

# Then try:
Get-GeoLocation
# This relies on the first of the four API features mentioned above. 
# If enabled it should tell you your current address, within 100 metres or so.

# Or try:
Get-GeoLocation -WithCoordinates
# This includes your current Lat/Lon (which comes in handy later for finding pizza and beer)

# How does googlemaps find your location? All you pass it is the mac address of the first two WIFI Access Points it can find, 
# via a line of code like this
netsh wlan show networks mode=Bssid | Where-Object{$_ -like "*BSSID*"} | %{($_.split(" ")[-1]).toupper()}
# Spooky eh? Let the consequences of that sink in.

# Now let's get some directions:
Get-Direction -from "250 Adelaide St, Brisbane" -to "30 Mary St, Brisbane" | ft -autosize
# This relies on the second of the four API features mentioned above.

# Example output:

<#
Instructions                                             Duration Distance Mode
------------                                             -------- -------- ----
Head southwest on Adelaide St toward Isles Ln            1 min    0.2 km   D...
Turn left onto Edward St                                 2 mins   0.4 km   D...
Turn right onto Mary St Destination will be on the right 2 mins   0.4 km   D...
#>

# Finally onto the fun one!

(Get-GeoLocation -WithCoordinates).Coordinates | Get-NearbyPlace -Radius 2000 -TypeOfPlace Restaurant -Keyword Pizza
# This relies on the 'places' api -- enabled at GooglePlaces_API_Key
# And tells you any pizza restaurants within a 2 kilometre radius. 
#If this returns no results then I suggest you move, because pizza is near the base of Maslow's hierarchy of needs.

(Get-GeoLocation -WithCoordinates).Coordinates | Get-NearbyPlace -Radius 2000 -TypeOfPlace Bar -Keyword Beer
#This is of similar if not greater importance. Your personal circumstances may of course vary.

Author: Prateek Singh