$User = "example@sharepoint.onmicrosoft.com"
$Password = Read-Host -Prompt "Please enter your password" 

$basePath = "C:\Temp"
$file = "ArchiveSites.csv"
$data = Import-Csv $file -Delimiter ';'

$totalRows = $data.Count
$processed = 0
$percent = $processed / $totalRows * 100

Write-Progress -Activity "Downloading sites" -PercentComplete $percent -CurrentOperation "Starting."

foreach ($site in $data) {
	$url = $site.SiteUrl
	$lf = $site.LocalFolder
	$folder = "$basePath\$lf"
	.\SPOnlineListDownloader.exe $url $user $password $folder
	$processed += 1 
	$percent = $processed / $totalRows * 100
	Write-Progress -Activity "Downloading sites" -PercentComplete $percent -CurrentOperation "Completed $url"
}