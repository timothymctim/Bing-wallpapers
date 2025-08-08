# Bing Wallpapers
# Fetch the Bing wallpaper image of the day
# <https://github.com/timothymctim/Bing-wallpapers>
#
# Copyright (c) 2015 Tim van de Kamp
# License: MIT license
Param(
    # Get the Bing image of this country
    [ValidateSet('auto', 'ar-XA', 'bg-BG', 'cs-CZ', 'da-DK', 'de-AT',
        'de-CH', 'de-DE', 'el-GR', 'en-AU', 'en-CA', 'en-GB', 'en-ID',
        'en-IE', 'en-IN', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-US',
        'en-XA', 'en-ZA', 'es-AR', 'es-CL', 'es-ES', 'es-MX', 'es-US',
        'es-XL', 'et-EE', 'fi-FI', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR',
        'he-IL', 'hr-HR', 'hu-HU', 'it-IT', 'ja-JP', 'ko-KR', 'lt-LT',
        'lv-LV', 'nb-NO', 'nl-BE', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT',
        'ro-RO', 'ru-RU', 'sk-SK', 'sl-SL', 'sv-SE', 'th-TH', 'tr-TR',
        'uk-UA', 'zh-CN', 'zh-HK', 'zh-TW')][string]$locale = 'auto',

    # Download the latest $files wallpapers.
    [ValidateRange("Positive")]
    [int]$files = 3,

    # Keep existing files
    [switch]$removeExistingFiles,

    # Resolution of the image to download
    [ValidateSet('auto', '800x600', '1024x768', '1280x720', '1280x768',
        '1366x768', '1920x1080', '1920x1200', '720x1280', '768x1024',
        '768x1280', '768x1366', '1080x1920')][string]$resolution = 'auto',

    # Destination folder to download the wallpapers to
    [string]$downloadFolder = $(Join-Path $([Environment]::GetFolderPath("MyPictures")) "Wallpapers"),

    # Use alternative API with > 500 images
    [switch]$useJsonSource
)
# Max item count: the number of images we'll query for
# [int]$maxItemCount = [System.Math]::max(1, [System.Math]::max($files, 8))
# URI to fetch the image locations from
if ($locale -eq 'auto') {
    $market = ""
}
else {
    $market = "&mkt=$locale"
}
[string]$hostname = "https://www.bing.com"
[string]$uri = "$hostname/HPImageArchive.aspx?format=xml$market&pid=hp"

# Get the appropiate screen resolution
if ($resolution -eq 'auto') {
    if ($PSVersionTable.OS.Contains('Windows')) {
        Write-Verbose 'On Windows'
        Add-Type -AssemblyName System.Windows.Forms
        $primaryScreen = [System.Windows.Forms.Screen]::AllScreens | Where-Object { $_.Primary -eq 'True' }
        $width = $primaryScreen.Bounds.Width
        $height = $primaryScreen.Bounds.Height
    }
    elseif ($PSVersionTable.OS.Contains('Darwin')) {
        Write-Verbose 'On MacOS'
        $null = system_profiler SPDisplaysDataType -json | ConvertFrom-Json -OutVariable displayInfo
        foreach ($item in $displayInfo) {
            if ($item[0].SPDisplaysDataType[0].spdisplays_ndrvs.spdisplays_main -eq 'spdisplays_yes') {
                [int]$width = $item[0].SPDisplaysDataType[0].spdisplays_ndrvs.spdisplays_resolution.split("@")[0].split('x')[0].Trim()
                [int]$height = $item[0].SPDisplaysDataType[0].spdisplays_ndrvs.spdisplays_resolution.split("@")[0].split('x')[1].Trim()
            }
        }
    }
    else {
        # Default for Linux
        $width = 1920
        $height = 1080
    }

    # Determine the resolution to download based on width and height
    if ($width -le 1024) {
        $resolution = '1024x768'
    }
    elseif ($width -le 1280) {
        $resolution = '1280x720'
    }
    elseif ($width -le 1366) {
        $resolution = '1366x768'
    }
    elseif ($height -le 1080) {
        $resolution = '1920x1080'
    }
    else {
        $resolution = '1920x1200'
    }
}

# Check if download folder exists and otherwise create it
if (!(Test-Path $downloadFolder)) {
    New-Item -ItemType Directory $downloadFolder
}

# Add paging support for when number requested > 8
# &idx=0&n=$maxItemCount
$pageSize = 8
$items = New-Object System.Collections.ArrayList
if ($files -gt 0) {
    if ($useJsonSource) {
        $jsonImgs = ConvertFrom-Json -InputObject $(Invoke-WebRequest -Uri 'https://api45gabs.azurewebsites.net/api/sample/bingphotos')
        for ($i = 0; $i -lt $files; $i++) {
            [datetime]$imageDate = [datetime]::ParseExact($jsonImgs[$i].startdate, 'yyyyMMdd', $null)
            [string]$imageUrl = "$hostname$($jsonImgs[$i].urlBase)_$resolution.jpg"
            # Add item to our array list
            $item = New-Object System.Object
            $item | Add-Member -Type NoteProperty -Name date -Value $imageDate
            $item | Add-Member -Type NoteProperty -Name url -Value $imageUrl
            $null = $items.Add($item)
        }
    }
    else {
        <# Action when all if and elseif conditions are false #>
        if ($files -gt 15) { $files = 15; Write-Debug "Bing API supports a maximum of 15 images" }
        $pageAndRemainder = [System.Math]::DivRem($files, $pageSize)
        # To support remainder logic vs 0 offset pages
        if (($pageAndRemainder.item2 -eq 0) -and ($pagesAndRemainder.item1 -ne 0)) {
            $pages = $pageAndRemainder.Item1 - 1
        }
        else {
            $pages = $pageAndRemainder.item1
        }
        for ($i = 0; $i -le $pages; $i++) {
            if ($pages -eq $i) { $pageItems = $pageAndRemainder.Item2 } else { $pageItems = $pageSize }
            [string]$pageUri = $uri + "&idx=$($i * $pageSize)&n=$pageItems"
            $request = Invoke-WebRequest -Uri $pageUri -DisableKeepAlive -UserAgent 'parp-1.0'
            [xml]$content = $request.Content
            
            foreach ($xmlImage in $content.images.image) {
                [datetime]$imageDate = [datetime]::ParseExact($xmlImage.startdate, 'yyyyMMdd', $null)
                [string]$imageUrl = "$hostname$($xmlImage.urlBase)_$resolution.jpg"
                
                # Add item to our array list
                $item = New-Object System.Object
                $item | Add-Member -Type NoteProperty -Name date -Value $imageDate
                $item | Add-Member -Type NoteProperty -Name url -Value $imageUrl
                $null = $items.Add($item)
            }
        }
    }
}

$items = $items | Sort-Object -Property date -Unique
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

if ($removeExistingFiles -and ($files -gt 0)) {
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
