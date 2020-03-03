Import-Module .\Wallpaper.ps1 -Force

Function Get-Rating {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]
        $FullName
    )

    begin {
        $ranking = 19 # extended attribute id "Rating"
        $shell = New-Object -COMObject Shell.Application
    }

    process {
        $path = $FullName

        $folder = Split-Path $path
        $file = Split-Path $path -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)
        
        Write-Progress -Activity "Get Ratings" -Status "Get rating attribute of $($path)" 
        $rating = $shellfolder.GetDetailsOf($shellfile, $ranking)

        $r = @{
            FullName = $path
            Rating   = $rating
        }
        return $r
    }
  
}

Function Get-FavoriteImages {
    Param([string]$Folder = (Get-Item -Path ".\").FullName    )

    $imgsRating = Get-ChildItem -Path $Folder -Filter *.jpg -Recurse | Get-Rating 
    $imgsRatingFourStars = $imgsRating | Where-Object { $_.Rating -like '4 Stars' }
    return $imgsRatingFourStars
}

Function Set-FavoriteBackground {
    Param([string]$Folder = (Get-Item -Path ".\").FullName    )

    $imgs = Get-FavoriteImages $Folder
    $r = Get-Random -Maximum ($imgs.Length - 1)
    $w = $imgs[$r].FullName

    Write-Host 'Setting Wallpaper to ' $($w)
    Set-Wallpaper -Path $w -Style Fill
}

Set-FavoriteBackground -Folder c:\life