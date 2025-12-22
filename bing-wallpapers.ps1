# Bing Wallpapers
# Fetch the Bing wallpaper image of the day
# <https://github.com/timothymctim/Bing-wallpapers>
#
# Copyright (c) 2015 Tim van de Kamp
# License: MIT license
Param(
    # Get the Bing image of this country
    [ValidateSet(
        'auto', 'ar-XA', 'bg-BG', 'cs-CZ', 'da-DK', 'de-AT', 'de-CH',
        'de-DE', 'el-GR', 'en-AU', 'en-CA', 'en-GB', 'en-ID', 'en-IE',
        'en-IN', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-US', 'en-XA',
        'en-ZA', 'es-AR', 'es-CL', 'es-ES', 'es-MX', 'es-US', 'es-XL',
        'et-EE', 'fi-FI', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR', 'he-IL',
        'hr-HR', 'hu-HU', 'it-IT', 'ja-JP', 'ko-KR', 'lt-LT', 'lv-LV',
        'nb-NO', 'nl-BE', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT', 'ro-RO',
        'ru-RU', 'sk-SK', 'sl-SL', 'sv-SE', 'th-TH', 'tr-TR', 'uk-UA',
        'zh-CN', 'zh-HK', 'zh-TW'
    )][string]$locale = 'auto',

    # Download the latest $files wallpapers
    [int]$files = 3,

    # Resolution of the image to download
    [ValidateSet(
        'auto', '800x600', '1024x768', '1280x720', '1280x768', '1366x768',
        '1920x1080', '1920x1200', '720x1280', '768x1024', '768x1280',
        '768x1366', '1080x1920', 'UHD'
    )][string]$resolution = 'auto',

    # Destination folder to download the wallpapers to
    [string]$downloadFolder = $(Join-Path $([Environment]::GetFolderPath("MyPictures")) "Wallpapers")
)

# Max item count: the number of images we'll query for
[int]$maxItemCount = [System.Math]::max(1, [System.Math]::max($files, 8))
# URI to fetch the image locations from
if ($locale -eq 'auto') {
    $market = ""
}
else {
    $market = "&mkt=$locale"
}
[string]$hostname = "https://www.bing.com"
[string]$uri = "$hostname/HPImageArchive.aspx?format=xml&idx=0&n=$maxItemCount$market"

# Get the appropiate screen resolution
if ($resolution -eq 'auto') {
    Add-Type -AssemblyName System.Windows.Forms
    $primaryScreen = [System.Windows.Forms.Screen]::AllScreens | Where-Object { $_.Primary -eq 'True' }
    if ($primaryScreen.Bounds.Width -le 1024) {
        $resolution = '1024x768'
    }
    elseif ($primaryScreen.Bounds.Width -le 1280) {
        $resolution = '1280x720'
    }
    elseif ($primaryScreen.Bounds.Width -le 1366) {
        $resolution = '1366x768'
    }
    elseif ($primaryScreen.Bounds.Height -le 1080) {
        $resolution = '1920x1080'
    }
    elseif ($primaryScreen.Bounds.Width -le 1920) {
        $resolution = '1920x1080'
    }
    else {
        $resolution = 'UHD'
    }
}

# Check if download folder exists and otherwise create it
if (!(Test-Path $downloadFolder)) {
    New-Item -ItemType Directory $downloadFolder
}

$request = Invoke-WebRequest -Uri $uri
[xml]$content = $request.Content

$items = New-Object System.Collections.ArrayList
foreach ($xmlImage in $content.images.image) {
    [datetime]$imageDate = [datetime]::ParseExact($xmlImage.startdate, 'yyyyMMdd', $null)
    [string]$imageUrl = "$hostname$($xmlImage.urlBase)_$resolution.jpg"

    # Add item to our array list
    $item = New-Object System.Object
    $item | Add-Member -Type NoteProperty -Name date -Value $imageDate
    $item | Add-Member -Type NoteProperty -Name url -Value $imageUrl
    $null = $items.Add($item)
}

# Keep only the most recent $files items to download
if (!($files -eq 0) -and ($items.Count -gt $files)) {
    # We have too many matches, keep only the most recent
    $items = $items | Sort-Object date
    while ($items.Count -gt $files) {
        # Pop the oldest item of the array
        $null, $items = $items
    }
}

Write-Host "Downloading images..."
$client = New-Object System.Net.WebClient
foreach ($item in $items) {
    $baseName = $item.date.ToString("yyyy-MM-dd")
    $destination = Join-Path $downloadFolder "$baseName.jpg"
    $url = $item.url

    # Download the enclosure if we haven't done so already
    if (!(Test-Path $destination)) {
        Write-Debug "Downloading image to $destination"
        $client.DownloadFile($url, "$destination")
    }
}

if ($files -gt 0) {
    # We do not want to keep every file; remove the old ones
    Write-Host "Cleaning the directory..."
    $i = 1
    Get-ChildItem -Filter "????-??-??.jpg" $downloadFolder | Sort-Object -Descending FullName | ForEach-Object {
        if ($i -gt $files) {
            # We have more files than we want, delete the extra files
            $fileName = $_.FullName
            Write-Debug "Removing file $fileName"
            Remove-Item "$fileName"
        }
        $i++
    }
}
