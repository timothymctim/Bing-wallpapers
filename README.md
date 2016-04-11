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
The script supports several option which allows you to customize its
behavior.

* `-locale` Get the Bing image of the day for this area.

  **Possible values** `'de-DE'`, `'en-AU'`, `'en-CA'`, `'en-NZ'`,
  `'en-UK'`, `'en-US'`, `'ja-JP'`, `'zh-CN'`

  **Default value** `'en-US'`

* `-files` Keep only this number of images in the folder, *any other
  file matching* `????-??-??.jpg` *will be* **removed**!

  **Default value** `3`

  **Remarks** Setting this option to `0` will keep all images and will
  not remove any file.

* `-resolution` Determines which image resolution will be downloaded.
  If set to `'auto'` the script will try to determine which resolution
  is more appropriate based on your primary screen resolution.

  **Possible values** `'auto'`, `'1024x768'`, `'1280x720'`,
  `'1366x768'`, `'1920x1080'`

  **Default value** `'auto'`

* `-downloadFolder` Destination folder to download the wallpapers to.

  **Default value**
  `"$([Environment]::GetFolderPath("MyPictures"))\Wallpapers"`
  (the subfolder `Wallpapers` inside your default Pictures folder)

  **Remarks** The folder will automatically be created if it doesn't
  exists already.

* `-proxy` proxy address.

  **Default value** `'white space'`
  

Example  
=========

ps:> .\bing-wallpapers.ps1 -proxy "http://10.10.2.3:2020"  
    

Set as your wallpaper
=====================
With a few additional steps you're able to automatically download the
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
Note that the script itself doesn't need to be run as administrator!

You can configure to run the script periodically using 'Task Scheduler'.
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
the images to (defaults to the folder `Wallpapers` inside your Pictures
folder).
