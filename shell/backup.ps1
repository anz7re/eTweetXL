$TargLoc = Join-Path -Path $HOME -ChildPath "\.z7\autokit\etweetxl\mtsett\mytarg.mt" 
$DestLoc = Join-Path -Path $HOME -ChildPath "\.z7\autokit\etweetxl\mtsett\mydest.mt" 
$X = 0

foreach($l in Get-Content $TargLoc){
if (Test-Path $TargLoc -PathType Leaf){
$TargStr = Get-Content $TargLoc | Select -Index $X}

if (Test-Path $DestLoc -PathType Leaf){
$DestStr = Get-Content $DestLoc | Select -Index $X}

Compress-Archive -Path $TargStr -DestinationPath $DestStr -CompressionLevel Optimal -Force
$X+=1
}

Remove-Item $TargLoc
Remove-Item $DestLoc
