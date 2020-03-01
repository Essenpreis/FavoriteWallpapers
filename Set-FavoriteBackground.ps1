Import-Module .\Wallpaper.ps1

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
        Write-Host 'Get attribs of ' $path

        $folder = Split-Path $path
        $file = Split-Path $path -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)

        $r = @{
            FullName = $path
            Rating   = $shellfolder.GetDetailsOf($shellfile, $ranking)
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
    Write-Host 'Setting Wallpaper to ' $imgs[$r].FullName

    Set-WallPaper -Path $imgs[$r].FullName -Style Fill
}

Set-FavoriteBackground -Folder c:\life