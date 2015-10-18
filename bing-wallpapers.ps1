# Bing Wallpapers
# Fetch the Bing wallpaper image of the day 
# <https://github.com/timothymctim/Bing-wallpapers>
#
# Copyright (c) 2015 Tim van de Kamp
# Licensed under the MIT license
Param(
	# Get the Bing image of this country
	# Possible values: "United States", "United Kingdom", "New Zealand", "Japan", "China", "Australia"
	[string]$country = "United States",

	# Download the latest $files wallpapers
	[int]$files = 3,

	# Destination folder to download the wallpapers to
	[string]$downloadFolder = "$([Environment]::GetFolderPath("MyPictures"))\Wallpapers",

	# Feed name
	[string]$feedName = "Bing Images"
)
# Feed URL
$feedUrl = "http://feeds.feedburner.com/bingimages"
# Max item count: download at least 60 feed items
$maxItemCount = [System.Math]::max($files * 6, 60)

# Check if download folder exists
if (!(Test-Path $downloadFolder)) {
	New-Item -ItemType Directory $downloadFolder
}

$fm = New-Object -ComObject "Microsoft.FeedsManager"
# Check if the feed exists; add feed if non-existing
if ($fm.RootFolder.ExistsFeed($feedName)) {
	$feed = $fm.RootFolder.GetFeed($feedName)
} else {
	$feed = $fm.RootFolder.CreateFeed($feedName, $feedUrl)
	$feed.DownloadEnclosuresAutomatically = $false
}

# Only update the feed if we haven't updated it in one day or the maxItemCount has increased
if (($feed.LastDownloadTime.AddDays(1).CompareTo([System.DateTime]::Now) -le 0) -or ($feed.MaxItemCount -lt $maxItemCount)) {
	Write-Host "Updating feed..."
	$feed.MaxItemCount = $maxItemCount
	$feed.Download()
}

$client = New-Object System.Net.WebClient
$itemMatches = New-Object System.Collections.ArrayList
foreach ($item in $feed.Items) {
	$title = $item.Title
	$url = $item.Enclosure.Url
	if ($url -and $title -Match $country) {
		# We've found a match but we're not sure if the wallpaper is recent enough
		$null = $itemMatches.Add($item)
	}
}

if (!($files -eq 0) -and ($itemMatches.Count -gt $files)) {
	# We have too many matches, keep only the most recent
	$itemMatches = $itemMatches|Sort PubDate
	while ($itemMatches.Count -gt $files) {
		# Pop the oldest item of the array
		$null, $itemMatches = $itemMatches
	}
}

Write-Host "Downloading enclosures..."
foreach ($item in $itemMatches) {
	$baseName = $item.Modified.ToString("yyyy-MM-dd")
	$destination = "$downloadFolder\$baseName.jpg"
	$url = $item.Enclosure.Url

	# Download the enclosure if we haven't done so already
	if (!(Test-Path $destination)) {
		Write-Debug "Downloading enclosure to $destination"
		$client.DownloadFile($url, "$destination")
	}
}

if ($files -gt 0) {
	# We do not want to keep every file; remove the old ones
	Write-Host "Cleaning the directory..."
	$i = 1
	Get-ChildItem -Filter "????-??-??.jpg" $downloadFolder | Sort -Descending FullName | ForEach-Object {
		$i++
		if ($i -gt $files) {
			# We have more files than we want, delete the extra files
			$fileName = $_.FullName
			Write-Debug "Removing file $fileName"
			Remove-Item "$fileName"
		}
	}
}
