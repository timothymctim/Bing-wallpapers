Bing image of the day
=====================
This Windows PowerShell script automatically fetches the Bing image of
the day.
Using this script you can set the Bing image of the day as your
wallpaper.

The script uses the RSS feed of
[I started something](http://www.istartedsomething.com/bingimages/) to
download the images.
It adds the feed to the Internet Explorer feed reader.
Together with a few extra steps, this script allows you to set your
wallpaper to the Bing image of the day, just like using [Bing
desktop](http://blogs.msdn.com/b/buckh/archive/2013/01/02/bing-desktop-set-your-background-to-the-bing-image-of-the-day.aspx)
(which you might not want to install or is unavailable in your region).

Script options
--------------
The script supports several option which allows you to customize its
behavior.

* `-country`
  Get the Bing image of the day for this country

  **Possible values** `"United States"`, `"United Kingdom"`,
  `"New Zealand"`, `"Japan"`, `"China"`, `"Australia"`

  **Default value** `"United States"`

* `-files` Keep only this number of images in the folder, *any other
  file matching* `????-??-??.jpg` *will be* **removed**!

  **Default value** `3`

  **Remarks** Setting this option to `0` will keep all images and will
  not remove any file.

* `-downloadFolder` Destination folder to download the wallpapers to

  **Default value**
  `"$([Environment]::GetFolderPath("MyPictures"))\Wallpapers"`
  (the subfolder `Wallpapers` inside your default Pictures folder)

  **Remarks** The folder will automatically be created if it doesn't
  exists already.

* `-feedName` The name of the Internet Explorer feed.

  **Default value** `"Bing Images"`

  **Remarks** The feed will automatically be created with this name if
  it doesn't exists already.

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
`Slideshow` as the `Background` type .
Hit the `Browse` button to select the folder you automatically download
the images to (defaults to the folder `Wallpapers` inside your Pictures
folder).
