#
#    Script copies all windows background and lock screen images from
#    Local Windows CDM folder to a Temp folder on C: drive called "Images"
#    and changes the file extensions to .png.
#    =)


# Balloon pop up stuff.
Add-Type -AssemblyName System.Windows.Forms 

$global:balloon = New-Object System.Windows.Forms.NotifyIcon

$path = (Get-Process -id $pid).Path

$ImagesFolder = "C:\Temp\Images"
$CDM_Folder = $home + "\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"

$added = "Images have been added  =)"
$not_added = "There are no new Images  =("

function OpenImagesFolder
{
    Invoke-Item $ImagesFolder
}

function Balloon-PopUp
{
    
    param(
        $text
    )

    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $balloon.BalloonTipText = "$text"
    $balloon.BalloonTipTitle = "$ImagesFolder" 
    $balloon.Visible = $true 
    $balloon.ShowBalloonTip(5000)
    # $balloon.BalloonTipClicked += [System.EventHandler](Balloon-Clicked)
    

}




# check if C:\Temp\Images folder already exists
if (Test-Path -Path $ImagesFolder)
{
    Write-Host "## '$ImagesFolder' already exists!!"
}
else
{
# if it doesn't create it =)
    Write-Host ""
    Write-Host "## Creating '$ImagesFolder'"
    New-Item $ImagesFolder -ItemType Directory
}


$files = Get-ChildItem $CDM_Folder

$flag = 0

foreach ($file in $files)
{
    
    # create this just to check Imagesfolder for existing image.
    $newName = $file.BaseName + '.png'

    # Write-Host $newName

    if (!(Test-Path $ImagesFolder\$newName))
    {
        # doesn't exist.
        Copy-Item $CDM_Folder\$file $ImagesFolder\$file
        Get-Item $ImagesFolder\$file | Rename-Item -NewName { [System.IO.Path]::ChangeExtension($_.Name, ".png") }
        $flag = 1
    }

}


# pop-up tells user if images have been added

$wshell = New-Object -ComObject Wscript.Shell

if (!$flag -eq 0)
{
#    $Output = $wshell.Popup($added)

    Balloon-PopUp $added

    # open Images Folder.
    OpenImagesFolder

}
else
{
#    $Output = $wshell.Popup($not_added)
    Balloon-PopUp $not_added
}






