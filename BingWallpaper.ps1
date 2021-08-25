###---------------------------------------------------------------
### Filename:           BingWallpaper.ps1
### Author:             David Rodgers
### Date:               2021.02.09
### Usage:              Gets/Saves/Sets Big Wallpaper of the Day
###---------------------------------------------------------------
### The [Microsoft Bing](bing.com) search engine provides a beautiful 
### picture every day, and I got tired of manually downloading and 
### renaming the picture each time I wanted to add one to my collection.
### I decided to solve the problem ### and cobbled together this 
### all-in-one solution. I run it as a nightly task. Problem solved.
###
### Using the Bing API, you can easily get the images for your own use.
### This script will grab the current Bing wallpaper of the day, save it 
### to a specified folder using the descriptive copyright title and date,
### and then set it ### as your desktop wallpaper. There are also options
### to specify the wallaper style (e.g. - Fill, Fit, Stretch, etc.) to use.
###---------------------------------------------------------------

###---------------------------------------------------------------
### Running as a scheduled task.
### Program/Script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
### Args: -ExecutionPolicy Bypass -File "C:\Scripts\BingWallPaper.ps1"
### Args: -ExecutionPolicy Bypass -File "D:\Work\Project\BingWallpaper\BingWallpaper.ps1"
### Args: -ExecutionPolicy Unrestricted -File "D:\Work\Project\BingWallpaper\BingWallpaper.ps1"

###---------------------------------------------------------------

###---------------------------------------------------------------
### Joe Espitia
### https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/
###---------------------------------------------------------------
Function Set-WallPaper {
 
    <#
     
        .SYNOPSIS
        Applies a specified wallpaper to the current user's desktop
        
        .PARAMETER Image
        Provide the exact path to the image
     
        .PARAMETER Style
        Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)
      
        .EXAMPLE
        Set-WallPaper -Image "C:\Wallpaper\Default.jpg"
        Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
      
    #>
     
    param (
        [parameter(Mandatory = $True)]
        # Provide path to image
        [string]$Image,
        # Provide wallpaper style that you would like applied
        [parameter(Mandatory = $False)]
        [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
        [string]$Style
    )
     
    $WallpaperStyle = Switch ($Style) {
      
        "Fill" { "10" }
        "Fit" { "6" }
        "Stretch" { "2" }
        "Tile" { "0" }
        "Center" { "0" }
        "Span" { "22" }
      
    }
     
    If ($Style -eq "Tile") {
     
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force
     
    }
    Else {
     
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force
     
    }
     
    Add-Type -TypeDefinition @" 
    using System; 
    using System.Runtime.InteropServices;
      
    public class Params
    { 
        [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
        public static extern int SystemParametersInfo (Int32 uAction, 
                                                       Int32 uParam, 
                                                       String lpvParam, 
                                                       Int32 fuWinIni);
    }
"@ 
      
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
      
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
      
    $ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
}

###---------------------------------------------------------------
### Martina Grom - @magrom
### https://blog.atwork.at/post/Use-the-Daily-Bing-picture-in-Teams-calls
###---------------------------------------------------------------
### Use the Bing.com API.
### The idx parameter determines the day: 0 is the current day, 1
### is the previous day, etc. This goes back for max 7 days. 
###---------------------------------------------------------------
### The n parameter defines how many pictures you want to load. 
### Usually, n=1 to get the latest picture (of today) only. 
### The mkt parameter defines the culture, like en-US, de-DE, etc.
###---------------------------------------------------------------

# API
$uri = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"

# Get the picture metadata
$response = Invoke-WebRequest -Method Get -Uri $uri

# Extract the image content
$body = ConvertFrom-Json -InputObject $response.Content
$fileurl = "https://www.bing.com/" + $body.images[0].url

# 2021.08.25: Try to get the UHD version if available
$UHD_fileurl = $fileurl.Replace("1920x1080","UHD")

# Determine filename for both HD & UHD images
$filename = $body.images[0].copyright.Split('-(', 2)[-2].Replace(" ", "-").Replace("?", "").Replace("-", " ").TrimEnd(' ') + " - " + $body.images[0].startdate + "_HD.jpg"

$UHD_filename = $body.images[0].copyright.Split('-(', 2)[-2].Replace(" ", "-").Replace("?", "").Replace("-", " ").TrimEnd(' ') + " - " + $body.images[0].startdate + "_UHD.jpg"

# Download the images to a specified folder
# $filepath = $PSScriptRoot+"\"+$filename
$filepath = "C:\Users\David\OneDrive\PhotoStream\Wallpaper\Bing\" + $filename
$UHD_filepath = "C:\Users\David\OneDrive\PhotoStream\Wallpaper\Bing\" + $UHD_filename
Invoke-WebRequest -Method Get -Uri $fileurl -OutFile $filepath
Invoke-WebRequest -Method Get -Uri $UHD_fileurl -OutFile $UHD_filepath

# Show the generated picture filepath
$filepath

# Use: Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
# Styles: Fill, Fit, Stretch, Tile, Center, Span
Set-WallPaper -Image "$filepath" -Style Fill

Exit
# END OF LINE