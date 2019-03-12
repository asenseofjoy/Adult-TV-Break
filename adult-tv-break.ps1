function getTotalSeconds ($folder, $file) {
   $LengthColumn = 27
   $objShell = New-Object -ComObject Shell.Application 
   $objFolder = $objShell.Namespace($Folder)
   $objFile = $objFolder.ParseName($File)
   $Length = $objFolder.GetDetailsOf($objFile, $LengthColumn)
   $timespan = [TimeSpan]::Parse($Length)
   $totalSeconds = $timespan.TotalSeconds
   $totalSeconds
}

function vibrate ($strength) {
   $password = "<password>" | ConvertTo-SecureString -asPlainText -Force
   $credentials = New-Object System.Management.Automation.PSCredential("admin",$password)
   $url = "http://192.168.1.68:25105/3?0262<6 letter ID of the Insteon plug>"
   $strengthHex = [Convert]::ToString($strength, 16)
   if($strength -eq 0) {
       $url += "0F1300=I=3"
   } else {
       $url += "0F11" + $strengthHex + "=I=3"
   }

   $url
   Invoke-WebRequest $url -Credential $credentials | Out-Null
}

function playVideo ($folder, $video) {
   &'C:\Program Files (x86)\VideoLAN\VLC\vlc.exe' ($folder + "\" + $video) -f
}

function playVideo ($folder, $video, $start) {
   &'C:\Program Files (x86)\VideoLAN\VLC\vlc.exe' ($folder + "\" + $video) --start-time $start -f
}

function playVideo ($folder, $video, $start, $stop) {
   &'C:\Program Files (x86)\VideoLAN\VLC\vlc.exe' ($folder + "\" + $video) --start-time $start --stop-time $stop -f
}

$videosFolder = "C:\VLC\Videos\"
$videos = Get-ChildItem $videosFolder

$commercialsFolder = "C:\VLC\Commercials\"
$commercials = Get-ChildItem $commercialsFolder
$commercialPos = 0

foreach($video in $videos) {

   $totalSeconds = getTotalSeconds $videosFolder $video
   $halfOfVideo = $totalSeconds/2

   #first half of show
   vibrate 0
   playVideo $videosFolder $video 1 $halfOfVideo
   Start-Sleep -Seconds $halfOfVideo

   #porn
   vibrate 255
   playVideo $commercialsFolder $commercials[$commercialPos]
   $commercialTime = getTotalSeconds $commercialsFolder $commercials[$commercialPos]
   Start-Sleep -Seconds $commercialTime
   $commercialPos++

   #second half of show
   vibrate 0
   playVideo $videosFolder $video $halfOfVideo
   Start-Sleep -Seconds $halfOfVideo

   #porn
   vibrate 255
   playVideo $commercialsFolder $commercials[$commercialPos]
   $commercialTime = getTotalSeconds $commercialsFolder $commercials[$commercialPos]
   Start-Sleep -Seconds $commercialTime
}
