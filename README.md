Bing image of the day
=====================
This Windows PowerShell script automatically fetches the Bing image of
the day.
Using this script you can set the Bing image of the day as your
wallpaper.

The script uses the XML page of [Microsoft Bing](https://www.bing.com/)
to download the images.
With a few extra steps, the script allows you to set your wallpaper to
the Bing image of the day, just like using [Bing
desktop](http://blogs.msdn.com/b/buckh/archive/2013/01/02/bing-desktop-set-your-background-to-the-bing-image-of-the-day.aspx)
(which might be unavailable in your region or you do not want to
install).

Script options
--------------
The script supports several options which allows you to customize the
behavior.

* `-locale` Get the Bing image of the day for this
  [region](https://msdn.microsoft.com/en-us/library/dd251064.aspx).

  **Possible values** `'auto'`, `'ar-XA'`, `'bg-BG'`, `'cs-CZ'`,
  `'da-DK'`, `'de-AT'`, `'de-CH'`, `'de-DE'`, `'el-GR'`, `'en-AU'`,
  `'en-CA'`, `'en-GB'`, `'en-ID'`, `'en-IE'`, `'en-IN'`, `'en-MY'`,
  `'en-NZ'`, `'en-PH'`, `'en-SG'`, `'en-US'`, `'en-XA'`, `'en-ZA'`,
  `'es-AR'`, `'es-CL'`, `'es-ES'`, `'es-MX'`, `'es-US'`, `'es-XL'`,
  `'et-EE'`, `'fi-FI'`, `'fr-BE'`, `'fr-CA'`, `'fr-CH'`, `'fr-FR'`,
  `'he-IL'`, `'hr-HR'`, `'hu-HU'`, `'it-IT'`, `'ja-JP'`, `'ko-KR'`,
  `'lt-LT'`, `'lv-LV'`, `'nb-NO'`, `'nl-BE'`, `'nl-NL'`, `'pl-PL'`,
  `'pt-BR'`, `'pt-PT'`, `'ro-RO'`, `'ru-RU'`, `'sk-SK'`, `'sl-SL'`,
  `'sv-SE'`, `'th-TH'`, `'tr-TR'`, `'uk-UA'`, `'zh-CN'`, `'zh-HK'`,
  `'zh-TW'`

  **Default value** `'auto'`

  **Remarks** By using the value `'auto'`, Bing will attempt to
  determine an applicable locale based on your IP address.

  Currently, only the values `'de-DE'`, `'en-AU'`, `'en-CA'`, `'en-GB'`,
  `'en-IN'`, `'en-US'`, `'fr-CA'`, `'fr-FR'`, `'ja-JP'`, and `'zh-CN'`
  will have their own localized version. Other values will be considered
  as the “Rest of the World” by Bing.

* `-files` Keep only this number of images in the folder, *any other
  file matching* `????-??-??.jpg` *will be* **removed**!

  **Default value** `3`

  **Remarks** Setting this option to `0` will keep all images and will
  not remove any file.

* `-resolution` Determines which image resolution will be downloaded.
  If set to `'auto'` the script will try to determine which resolution
  is more appropriate based on your primary screen resolution.

  **Possible values** `'auto'`, `'800x600'`, `'1024x768'`, `'1280x720'`,
  `'1280x768'`, `'1366x768'`, `'1920x1080'`, `'1920x1200'`, `'720x1280'`,
  `'768x1024'`, `'768x1280'`, `'768x1366'`, `'1080x1920'`, `'UHD'`

  **Default value** `'auto'`

* `-downloadFolder` Destination folder to download the wallpapers to.

  **Default value**
  `"$([Environment]::GetFolderPath("MyPictures"))\Wallpapers"`
  (the subfolder `Wallpapers` inside your default Pictures folder)

  **Remarks** The folder will automatically be created if it doesn’t
  exist already.

Set as your wallpaper
=====================
With a few additional steps you’re able to automatically download the
latest images and set them as your wallpaper.

Automatically run the script
----------------------------
First, make sure that you can actually run PowerShell scripts.
You might have to set the execution policy to unrestricted by running
`Set-ExecutionPolicy Unrestricted` in a PowerShell window executed with
administrator rights.
Additionally, you might need to unblock the file since you downloaded
the file from an untrusted source on the Internet.
You can do this by running `Unblock-File <path to the script>` as
administrator.
Note that the script itself doesn’t need to be run as administrator!

You can configure to run the script periodically using “Task Scheduler.”
Open Task Scheduler and click `Action` ⇨ `Create Task…`.
Enter a name and description that you like.
Next, add a trigger to run the task once a day.
Finally, add the script as an action.
Run the program `powershell` with the arguments `-WindowStyle Hidden
-file "<path to the script>" <optional script arguments>`.

Changing your background settings
---------------------------------
Go to `Settings` ⇨ `Personalization` ⇨ `Background` and select
`Slideshow` as the `Background` type.
Hit the `Browse` button to select the folder you automatically download
the images to (the default is the folder `Wallpapers` inside your
Pictures folder).
